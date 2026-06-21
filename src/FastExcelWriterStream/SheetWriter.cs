using System.Globalization;
using System.Text;

namespace FastExcelWriterStream;

internal sealed class SheetWriter : IDisposable
{
    private readonly Stream      _entryStream;
    private readonly StreamWriter _writer;
    private readonly SheetConfig  _sheet;
    private readonly int          _styleIndex;
    private readonly string       _styleAttr;

    // Cache column names - max 16384 columns in Excel
    private static readonly string[] ColNameCache = BuildColNameCache(50);

    private int  _lastRow = 0;
    private bool _rowOpen = false;
    private bool _disposed;
    internal bool IsDisposed => _disposed;
    private bool _headerWritten = false;

    // Auto-width tracking: col index → max char length seen
    private Dictionary<int, int>? _colMaxLen = null;
    internal bool AutoWidthEnabled { get; private set; } = false;



    private readonly bool _leaveStreamOpen;

    // buffer برای sheetData وقتی auto-width فعاله
    private MemoryStream? _dataBuffer = null;
    private StreamWriter?  _dataWriter = null;

    internal SheetWriter(Stream entryStream, SheetConfig sheet, int styleIndex, bool leaveStreamOpen = false)
    {
        _entryStream     = entryStream;
        _sheet           = sheet;
        _styleIndex      = styleIndex;
        _styleAttr       = styleIndex > 0 ? $" s=\"{styleIndex}\"" : "";
        _leaveStreamOpen = leaveStreamOpen;

        _writer = new StreamWriter(_entryStream, new UTF8Encoding(false), bufferSize: 1024 * 128, leaveOpen: true);
        // WriteHeader لیزی صدا زده میشه - بعد از EnableAutoWidth اگه لازم بود
    }

    private int _autoWidthMaxSampleRows = 0;  // 0 = همه ردیف‌ها
    private int _autoWidthSampledRows   = 0;  // تعداد ردیف‌های sample شده

    internal void EnableAutoWidth(int maxSampleRows = 0)
    {
        // اول header رو بنویس (بدون cols و بدون sheetData)
        if (!_headerWritten) WriteHeaderOnly();

        AutoWidthEnabled          = true;
        _colMaxLen                = new Dictionary<int, int>();
        _autoWidthMaxSampleRows   = maxSampleRows;
        _dataBuffer               = new MemoryStream(1024 * 256);
        _dataWriter               = new StreamWriter(_dataBuffer, new UTF8Encoding(false), bufferSize: 1024 * 64, leaveOpen: true);
        // sheetData به buffer نوشته میشه
        _dataWriter.Write("<sheetData>");
    }

    /// <summary>آیا هنوز باید طول ستون رو track کنیم؟</summary>
    private bool ShouldTrackWidth => _colMaxLen != null &&
        (_autoWidthMaxSampleRows == 0 || _autoWidthSampledRows <= _autoWidthMaxSampleRows);

    // ── Public API ──────────────────────────────────────────────────

    internal void Write(string value, int col, int row, DataType dataType = DataType.Text, int styleIndex = -1)
    {
        EnsureRow(row);
        WriteCell(col, row, value, dataType, styleIndex);
    }

    internal void WriteNumber(double value, int col, int row, int styleIndex = -1)
        => WriteNumberDirect(value, col, row, styleIndex);

    internal void WriteBool(bool value, int col, int row, int styleIndex = -1)
    {
        EnsureRow(row);
        var colName   = GetColName(col);
        var styleAttr = styleIndex >= 0
            ? (styleIndex > 0 ? $" s=\"{styleIndex}\"" : "")
            : _styleAttr;
        var w = _dataWriter ?? _writer;
        w.Write("<c r=\"");
        w.Write(colName);
        w.Write(row);
        w.Write("\" t=\"b\"");
        w.Write(styleAttr);
        w.Write("><v>");
        w.Write(value ? '1' : '0');
        w.Write("</v></c>");
    }

    internal void WriteFormula(string formula, int col, int row)
        => Write(formula, col, row, DataType.Formula);

    internal void WriteFormula(FormulaType formulaType, int col, int row, int dataCol, int dataRowStart, int dataRowEnd)
    {
        var range   = $"{GetColName(dataCol)}{dataRowStart}:{GetColName(dataCol)}{dataRowEnd}";
        var formula = $"{formulaType.ToString().ToUpperInvariant()}({range})";
        Write(formula, col, row, DataType.Formula);
    }

    // ── Internals ───────────────────────────────────────────────────

    private void EnsureRow(int row)
    {
        EnsureHeader();

        if (row < _lastRow)
            throw new InvalidOperationException(
                $"Rows must be written in ascending order. Last: {_lastRow}, Requested: {row}");

        if (row > _lastRow)
        {
            // اگه sample rows تموم شد، cols رو بنویس و به direct write برو
            if (_dataWriter != null &&
                _autoWidthMaxSampleRows > 0 &&
                _autoWidthSampledRows >= _autoWidthMaxSampleRows)
            {
                FlushSampleAndSwitchToDirect();
            }

            var w = _dataWriter ?? _writer;
            if (_rowOpen) { w.Write("</row>"); _rowOpen = false; }

            var ht = GetRowHeight(row);
            if (ht.HasValue)
            {
                w.Write("<row r=\"");
                w.Write(row);
                w.Write("\" ht=\"");
                w.Write(ht.Value.ToString(CultureInfo.InvariantCulture));
                w.Write("\" customHeight=\"1\">");
            }
            else
            {
                w.Write("<row r=\"");
                w.Write(row);
                w.Write("\">");
            }

            _lastRow = row;
            _rowOpen = true;
            _autoWidthSampledRows++;
        }
    }

    /// <summary>
    /// وقتی sample rows تموم شد:
    /// cols رو به stream اصلی بنویس، بعد buffer رو کپی کن، بعد مستقیم بنویس
    /// </summary>
    private void FlushSampleAndSwitchToDirect()
    {
        if (_dataWriter == null || _colMaxLen == null) return;

        // 1. بستن ردیف باز در buffer
        if (_rowOpen) { _dataWriter.Write("</row>"); _rowOpen = false; }
        _dataWriter.Flush();

        // 2. نوشتن cols به stream اصلی
        WriteCols();

        // 3. کپی buffer (sheetData تا اینجا) به stream اصلی
        // _writer.Flush() قبل از CopyTo به BaseStream - مهم!
        _writer.Flush();
        _dataBuffer!.Position = 0;
        _dataBuffer.CopyTo(_writer.BaseStream);

        // 4. آزاد کردن buffer
        _dataWriter.Dispose();
        _dataBuffer.Dispose();
        _dataWriter = null;
        _dataBuffer = null;
        _rowOpen    = false;
    }

    private void WriteCols()
    {
        if (_colMaxLen == null || _colMaxLen.Count == 0) return;
        _writer.Write("<cols>");
        foreach (var kv in _colMaxLen.OrderBy(x => x.Key))
        {
            var width = Math.Min(60, Math.Max(8, kv.Value * 1.0 + 1));
            _writer.Write(
                $"<col min=\"{kv.Key}\" max=\"{kv.Key}\" " +
                $"width=\"{width.ToString(CultureInfo.InvariantCulture)}\" " +
                $"bestFit=\"1\" customWidth=\"1\"/>");
        }
        _writer.Write("</cols>");
        _writer.Flush();
    }

    private void WriteCell(int col, int row, string? value, DataType dataType, int styleIndex = -1)
    {
        value ??= string.Empty;

        // track max length for auto-width
        if (ShouldTrackWidth)
        {
            var len = value.Length;
            if (!_colMaxLen!.TryGetValue(col, out var cur) || len > cur)
                _colMaxLen[col] = len;
        }
        var colName   = GetColName(col);
        var styleAttr = styleIndex >= 0
            ? (styleIndex > 0 ? $" s=\"{styleIndex}\"" : "")
            : _styleAttr;
        var w = _dataWriter ?? _writer;

        switch (dataType)
        {
            case DataType.Number:
                w.Write("<c r=\"");
                w.Write(colName);
                w.Write(row);
                w.Write('"');
                w.Write(styleAttr);
                w.Write("><v>");
                w.Write(value);
                w.Write("</v></c>");
                break;

            case DataType.Formula:
                w.Write("<c r=\"");
                w.Write(colName);
                w.Write(row);
                w.Write('"');
                w.Write(styleAttr);
                w.Write("><f>");
                w.Write(NeedsEscape(value) ? XmlHelper.EscapeXml(value) : value);
                w.Write("</f></c>");
                break;

            case DataType.Boolean:
                w.Write("<c r=\"");
                w.Write(colName);
                w.Write(row);
                w.Write("\" t=\"b\"");
                w.Write(styleAttr);
                w.Write("><v>");
                // قبول می‌کنیم: "1"/"0"/"true"/"false"
                w.Write(value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ? '1' : '0');
                w.Write("</v></c>");
                break;

            default:
                w.Write("<c r=\"");
                w.Write(colName);
                w.Write(row);
                w.Write("\" t=\"inlineStr\"");
                w.Write(styleAttr);
                w.Write("><is><t>");
                w.Write(NeedsEscape(value) ? XmlHelper.EscapeXml(value) : value);
                w.Write("</t></is></c>");
                break;
        }
    }

    // برای اعداد مستقیم بدون تبدیل به string
    private void WriteNumberDirect(double value, int col, int row, int styleIndex = -1)
    {
        EnsureRow(row);
        // track length for auto-width
        if (ShouldTrackWidth)
        {
            var len = value.ToString(CultureInfo.InvariantCulture).Length;
            if (!_colMaxLen!.TryGetValue(col, out var cur) || len > cur)
                _colMaxLen[col] = len;
        }
        var colName   = GetColName(col);
        var styleAttr = styleIndex >= 0
            ? (styleIndex > 0 ? $" s=\"{styleIndex}\"" : "")
            : _styleAttr;
        var w = _dataWriter ?? _writer;
        w.Write("<c r=\"");
        w.Write(colName);
        w.Write(row);
        w.Write('"');
        w.Write(styleAttr);
        w.Write("><v>");
        w.Write(value.ToString(CultureInfo.InvariantCulture));
        w.Write("</v></c>");
    }

    // فقط وقتی واقعاً نیاز باشه escape می‌کنیم
    private static bool NeedsEscape(string s)
    {
        foreach (var c in s)
            if (c == '&' || c == '<' || c == '>' || c == '"' || c == '\'' || XmlHelper.IsIllegalXmlChar(c))
                return true;
        return false;
    }

    private static string GetColName(int col)
        => col <= ColNameCache.Length ? ColNameCache[col - 1] : XmlHelper.ColName(col);

    private double? GetRowHeight(int row)
    {
        if (_sheet.RowsHeight == null || row - 1 >= _sheet.RowsHeight.Count) return null;
        return _sheet.RowsHeight[row - 1];
    }

    /// <summary>اگه header نوشته نشده، بنویس</summary>
    private void EnsureHeader()
    {
        if (_headerWritten) return;
        if (AutoWidthEnabled)
        {
            // EnableAutoWidth این کار رو کرده
            _headerWritten = true;
            return;
        }
        WriteHeader();
    }

    /// <summary>فقط worksheet tag و sheetViews - بدون cols و sheetData</summary>
    private void WriteHeaderOnly()
    {
        _writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        _writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        if (_sheet.FreezeRows > 0 || _sheet.FreezeCols > 0 || _sheet.RightToLeft)
        {
            var rtlAttr = _sheet.RightToLeft ? " rightToLeft=\"1\"" : "";
            _writer.Write($"<sheetViews><sheetView workbookViewId=\"0\"{rtlAttr}>");
            if (_sheet.FreezeRows > 0 || _sheet.FreezeCols > 0)
            {
                var topLeft = XmlHelper.ColName(_sheet.FreezeCols + 1) + (_sheet.FreezeRows + 1);
                _writer.Write(
                    $"<pane xSplit=\"{_sheet.FreezeCols}\" ySplit=\"{_sheet.FreezeRows}\" " +
                    $"topLeftCell=\"{topLeft}\" activePane=\"bottomRight\" state=\"frozen\"/>");
            }
            _writer.Write("</sheetView></sheetViews>");
        }
        _writer.Flush();
        _headerWritten = true;
    }

    private void WriteHeader()
    {
        _headerWritten = true;
        _writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        _writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        // ترتیب عناصر worksheet در OOXML اجباریه: sheetViews باید قبل از cols بیاد
        if (_sheet.FreezeRows > 0 || _sheet.FreezeCols > 0 || _sheet.RightToLeft)
        {
            var rtlAttr = _sheet.RightToLeft ? " rightToLeft=\"1\"" : "";
            _writer.Write($"<sheetViews><sheetView workbookViewId=\"0\"{rtlAttr}>");
            if (_sheet.FreezeRows > 0 || _sheet.FreezeCols > 0)
            {
                var topLeft = XmlHelper.ColName(_sheet.FreezeCols + 1) + (_sheet.FreezeRows + 1);
                _writer.Write(
                    $"<pane xSplit=\"{_sheet.FreezeCols}\" ySplit=\"{_sheet.FreezeRows}\" " +
                    $"topLeftCell=\"{topLeft}\" activePane=\"bottomRight\" state=\"frozen\"/>");
            }
            _writer.Write("</sheetView></sheetViews>");
        }

        // اگه auto-width فعاله، cols بعد از نوشتن داده‌ها در WriteFooter نوشته میشه
        if (!AutoWidthEnabled && _sheet.ColumnsWidth?.Count > 0)
        {
            _writer.Write("<cols>");
            for (int i = 0; i < _sheet.ColumnsWidth.Count; i++)
                _writer.Write(
                    $"<col min=\"{i+1}\" max=\"{i+1}\" width=\"{_sheet.ColumnsWidth[i].ToString(CultureInfo.InvariantCulture)}\" customWidth=\"1\"/>");
            _writer.Write("</cols>");
        }

        if (_dataWriter != null)
            _dataWriter.Write("<sheetData>");
        else
        {
            _writer.Write("<sheetData>");
            _writer.Flush();
        }
    }

    private void WriteFooter()
    {
        EnsureHeader();
        var w = _dataWriter ?? _writer;
        if (_rowOpen) { w.Write("</row>"); _rowOpen = false; }
        w.Write("</sheetData>");

        // اگه هنوز buffer فعاله (maxSampleRows نرسیده یا نامحدود)
        if (_dataWriter != null && _colMaxLen != null)
        {
            _dataWriter.Flush();
            WriteCols();
            _writer.Flush(); // مهم قبل از CopyTo
            _dataBuffer!.Position = 0;
            _dataBuffer.CopyTo(_writer.BaseStream);
            _dataWriter.Dispose();
            _dataBuffer.Dispose();
            _dataWriter = null;
            _dataBuffer = null;
        }

        if (_sheet.MergeRanges?.Count > 0)
        {
            _writer.Write("<mergeCells>");
            foreach (var m in _sheet.MergeRanges)
                _writer.Write(
                    $"<mergeCell ref=\"{XmlHelper.ColName(m.ColStart)}{m.RowStart}:" +
                    $"{XmlHelper.ColName(m.ColEnd)}{m.RowEnd}\"/>");
            _writer.Write("</mergeCells>");
        }

        if (_sheet.FilterRange != null)
        {
            var fr = _sheet.FilterRange;
            _writer.Write(
                $"<autoFilter ref=\"{XmlHelper.ColName(fr.ColStart)}{fr.RowStart}:" +
                $"{XmlHelper.ColName(fr.ColEnd)}{fr.RowEnd}\"/>");
        }

        _writer.Write("</worksheet>");
        _writer.Flush();
    }

    /// <summary>برگرداندن عرض محاسبه‌شده برای هر ستون</summary>
    internal Dictionary<int, double> GetAutoWidths()
    {
        var result = new Dictionary<int, double>();
        if (_colMaxLen == null) return result;
        foreach (var kv in _colMaxLen)
        {
            // تقریب: هر کاراکتر ~1.2 واحد عرض، حداقل 8، حداکثر 60
            var width = Math.Min(60, Math.Max(8, kv.Value * 1.0 + 1));
            result[kv.Key] = width;
        }
        return result;
    }

    private static string[] BuildColNameCache(int count)
    {
        var cache = new string[count];
        for (int i = 0; i < count; i++)
            cache[i] = XmlHelper.ColName(i + 1);
        return cache;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        WriteFooter();
        // همیشه StreamWriter رو dispose کن تا buffer کاملاً flush بشه
        // چون leaveOpen: true ست شده، stream زیرین بسته نمیشه
        _writer.Dispose();
        // اگه leaveStreamOpen نبود، stream رو هم ببند
        if (!_leaveStreamOpen)
            _entryStream.Dispose();
    }
}
