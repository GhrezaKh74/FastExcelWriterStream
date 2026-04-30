using System.Text;
using SharpCompress.Common;
using SharpCompress.Writers.Zip;

// alias برای جلوگیری از تداخل با SharpCompress.CompressionLevel
using DotNetCompressionLevel = System.IO.Compression.CompressionLevel;
using DotNetZipArchive       = System.IO.Compression.ZipArchive;
using DotNetZipArchiveMode   = System.IO.Compression.ZipArchiveMode;

namespace FastExcelWriter;

public sealed class ExcelWriter : IDisposable
{
    // ── ثابت‌ها ──────────────────────────────────────────────────────
    public const int MaxSheets = 10;
    public const int MaxRows   = 1_048_576;
    public const int MaxCols   = 50;

    // ── فیلدها ──────────────────────────────────────────────────────
    private readonly Stream             _outputStream;
    private readonly bool               _leaveOpen;
    private readonly List<SheetConfig>  _sheets;
    private readonly List<StyleConfig>  _styles = new();
    private readonly ExcelWriterOptions _options;

    // StreamingZip
    private DotNetZipArchive?          _zipArchive;
    private readonly List<(Stream stream, SheetWriter sw)> _streamingWriters = new();

    // SharpCompressZip
    private ZipWriter?                                        _zipWriter;
    private readonly List<(Stream? stream, SheetWriter sw)>  _sharpWriters = new();

    // Auto-split (AutoWrite)
    private int    _autoCurrentSheet    = 1;
    private int    _autoCurrentRow      = 1;
    private int    _autoCurrentCol      = 1;
    private int    _autoTotalCols       = 0;
    private string _autoSheetNamePrefix = "Sheet";

    // Auto sheet split (Write معمولی)
    private bool   _autoSheetSplitEnabled  = false;
    private string _autoSheetSplitPrefix   = "Sheet";
    private int      _autoSheetSplitCurrent   = 1;
    private string[]? _autoSheetSplitHeader   = null;
    private int       _autoSheetSplitFreezeRows = 0;
    private bool      _autoSheetSplitRtl        = false;

    private bool _disposed;

    // ── Constructors ────────────────────────────────────────────────

    public ExcelWriter(string filePath, SheetConfig? sheet = null, StyleConfig? style = null, ExcelWriterOptions? options = null)
        : this(OpenFile(filePath), leaveOpen: false, sheet, style, options) { }

    public ExcelWriter(Stream stream, bool leaveOpen = false, SheetConfig? sheet = null, StyleConfig? style = null, ExcelWriterOptions? options = null)
    {
        _outputStream = stream ?? throw new ArgumentNullException(nameof(stream));
        _leaveOpen    = leaveOpen;
        _options      = options ?? new ExcelWriterOptions();
        _sheets       = new List<SheetConfig> { sheet ?? new SheetConfig() };

        if (style != null) _styles.Add(style);

        if (_options.ZipMode == ZipMode.SharpCompressZip)
        {
            _zipWriter = new ZipWriter(
                _outputStream,
                new ZipWriterOptions(
                    CompressionType.Deflate,
                    SharpCompress.Compressors.Deflate.CompressionLevel.BestSpeed)
                {
                    LeaveStreamOpen = true,
                    UseZip64        = true
                });
        }
        else
        {
            _zipArchive = new DotNetZipArchive(_outputStream, DotNetZipArchiveMode.Create, leaveOpen: true);
        }

        OpenSheetWriter(0);
    }

    // ── Single-sheet API ────────────────────────────────────────────

    public void Write(string value, int col, int row, DataType dataType = DataType.Text, int styleIndex = -1)
    {
        ValidateCol(col);

        if (_autoSheetSplitEnabled)
        {
            var (sheetIdx, mappedRow) = ResolveAutoSheetSplit(row);
            GetWriter(sheetIdx).Write(value, col, mappedRow, dataType, styleIndex);
        }
        else
        {
            ValidateRow(row);
            GetWriter(1).Write(value, col, row, dataType, styleIndex);
        }
    }

    public void WriteNumber(double value, int col, int row, int styleIndex = -1)
    {
        ValidateCol(col);

        if (_autoSheetSplitEnabled)
        {
            var (sheetIdx, mappedRow) = ResolveAutoSheetSplit(row);
            GetWriter(sheetIdx).WriteNumber(value, col, mappedRow, styleIndex);
        }
        else
        {
            ValidateRow(row);
            GetWriter(1).WriteNumber(value, col, row, styleIndex);
        }
    }

    public void WriteFormula(string formula, int col, int row)
    {
        ValidateCol(col);
        ValidateRow(row);
        GetWriter(1).WriteFormula(formula, col, row);
    }

    public void WriteFormula(FormulaType formulaType, int col, int row, int dataCol, int dataRowStart, int dataRowEnd)
    {
        ValidateCol(col);
        ValidateRow(row);
        GetWriter(1).WriteFormula(formulaType, col, row, dataCol, dataRowStart, dataRowEnd);
    }

    // ── Multi-sheet API ─────────────────────────────────────────────

    public int AddSheet(SheetConfig sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (_sheets.Count >= MaxSheets)
            throw new InvalidOperationException($"ماکسیمم تعداد شیت ({MaxSheets}) پر شده است.");

        if (_options.ZipMode == ZipMode.StreamingZip)
        {
            var (prevStream, prevWriter) = _streamingWriters[^1];
            if (!prevWriter.IsDisposed) prevWriter.Dispose();
            prevStream.Dispose();
        }
        else // SharpCompressZip
        {
            var (prevStream, prevWriter) = _sharpWriters[^1];
            if (!prevWriter.IsDisposed) prevWriter.Dispose();
            prevStream?.Dispose();
        }

        _sheets.Add(sheet);
        OpenSheetWriter(_sheets.Count - 1);
        return _sheets.Count;
    }

    public void Write(string value, int col, int row, int sheetIndex, DataType dataType = DataType.Text, int styleIndex = -1)
    {
        ValidateCol(col);
        ValidateRow(row);
        GetWriter(sheetIndex).Write(value, col, row, dataType, styleIndex);
    }

    public void WriteNumber(double value, int col, int row, int sheetIndex, int styleIndex = -1)
    {
        ValidateCol(col);
        ValidateRow(row);
        GetWriter(sheetIndex).WriteNumber(value, col, row, styleIndex);
    }

    public void WriteFormula(string formula, int col, int row, int sheetIndex)
    {
        ValidateCol(col);
        ValidateRow(row);
        GetWriter(sheetIndex).WriteFormula(formula, col, row);
    }

    // ── Enable Auto Sheet Split ──────────────────────────────────────

    /// <summary>
    /// فعال کردن auto sheet split.
    /// بعد از فعال کردن، همون Write معمولی رو استفاده کن.
    /// وقتی row از 1,048,576 گذشت، خودکار شیت جدید باز میشه.
    /// </summary>
    public void EnableAutoSheetSplit(
        string sheetNamePrefix = "Sheet",
        string[]? headerRow    = null,
        int freezeRows         = 0,
        bool rightToLeft       = false)
    {
        _autoSheetSplitEnabled    = true;
        _autoSheetSplitPrefix     = sheetNamePrefix;
        _autoSheetSplitCurrent    = 1;
        _autoSheetSplitHeader     = headerRow;
        _autoSheetSplitFreezeRows = freezeRows;
        _autoSheetSplitRtl        = rightToLeft;
        _sheets[0].Name           = $"{sheetNamePrefix} 1";

        // freeze و RTL رو روی شیت اول اعمال کن
        if (freezeRows > 0)
            _sheets[0].FreezeRows = freezeRows;
        if (rightToLeft)
            _sheets[0].RightToLeft = true;
    }

    // ── Auto Width API ──────────────────────────────────────────────

    /// <summary>
    /// فعال کردن تنظیم خودکار عرض ستون‌ها بر اساس محتوا.
    /// باید قبل از هر Write صدا زده بشه.
    /// عرض هر ستون بر اساس طولانی‌ترین مقدار آن ستون محاسبه میشه.
    /// </summary>
    /// <summary>
    /// فعال کردن تنظیم خودکار عرض ستون‌ها.
    ///
    /// maxSampleRows:
    ///   0 = همه ردیف‌ها (RAM بیشتر)
    ///   N = فقط N ردیف اول sample میشه، بقیه مستقیم نوشته میشن (توصیه: 500-2000)
    ///
    /// مثال:
    ///   ew.EnableAutoWidth(1000); // فقط ۱۰۰۰ ردیف اول برای محاسبه عرض
    /// </summary>
    public void EnableAutoWidth(int maxSampleRows = 1000)
    {
        foreach (var (_, sw) in _streamingWriters)
            sw.EnableAutoWidth(maxSampleRows);
        foreach (var (_, sw) in _sharpWriters)
            sw.EnableAutoWidth(maxSampleRows);
    }

    // ── Style API ───────────────────────────────────────────────────

    /// <summary>
    /// اضافه کردن style سفارشی و برگرداندن index آن.
    /// از index برگشتی در Write/WriteNumber استفاده کن.
    ///
    /// مثال:
    ///   int amountStyle = ew.AddStyle(new StyleConfig { NumberFormat = NumberFormat.Thousands });
    ///   int dollarStyle = ew.AddStyle(new StyleConfig { NumberFormat = NumberFormat.ThousandsDecimal });
    ///   ew.WriteNumber(22954062, col, row, styleIndex: amountStyle);
    /// </summary>
    public int AddStyle(StyleConfig style)
    {
        ArgumentNullException.ThrowIfNull(style);
        _styles.Add(style);
        return _styles.Count; // 1-based (0 = no style)
    }

    // ── Auto-split API ──────────────────────────────────────────────

    /// <summary>
    /// شروع حالت auto-split.
    /// تعداد ستون‌ها رو مشخص کن، بعد AutoWrite صدا بزن.
    /// </summary>
    public void BeginAutoSplit(int cols, string sheetNamePrefix = "Sheet")
    {
        if (cols < 1 || cols > MaxCols)
            throw new ArgumentOutOfRangeException(nameof(cols),
                $"تعداد ستون باید بین 1 و {MaxCols} باشد.");

        _autoTotalCols       = cols;
        _autoCurrentRow      = 1;
        _autoCurrentCol      = 1;
        _autoCurrentSheet    = 1;
        _autoSheetNamePrefix = sheetNamePrefix;
        _sheets[0].Name      = $"{sheetNamePrefix} 1";
    }

    public void AutoWrite(string value, DataType dataType = DataType.Text)
    {
        if (_autoTotalCols == 0)
            throw new InvalidOperationException("ابتدا BeginAutoSplit() را فراخوانی کنید.");

        if (_autoCurrentRow > MaxRows) SplitToNextSheet();

        GetWriter(_autoCurrentSheet).Write(value, _autoCurrentCol, _autoCurrentRow, dataType);
        AdvanceAutoPosition();
    }

    public void AutoWriteNumber(double value)
    {
        if (_autoTotalCols == 0)
            throw new InvalidOperationException("ابتدا BeginAutoSplit() را فراخوانی کنید.");

        if (_autoCurrentRow > MaxRows) SplitToNextSheet();

        GetWriter(_autoCurrentSheet).WriteNumber(value, _autoCurrentCol, _autoCurrentRow);
        AdvanceAutoPosition();
    }

    private void SplitToNextSheet()
    {
        if (_autoCurrentSheet >= MaxSheets)
            throw new InvalidOperationException(
                $"ماکسیمم شیت ({MaxSheets}) پر شده است.");

        AddSheet(new SheetConfig { Name = $"{_autoSheetNamePrefix} {_autoCurrentSheet + 1}" });
        _autoCurrentSheet++;
        _autoCurrentRow = 1;
        _autoCurrentCol = 1;
    }

    private void AdvanceAutoPosition()
    {
        _autoCurrentCol++;
        if (_autoCurrentCol > _autoTotalCols)
        {
            _autoCurrentCol = 1;
            _autoCurrentRow++;
        }
    }

    // ── Auto Sheet Split Resolver ────────────────────────────────────

    private (int sheetIndex, int mappedRow) ResolveAutoSheetSplit(int row)
    {
        if (row <= MaxRows)
            return (_autoSheetSplitCurrent, row);

        int targetSheet = ((row - 1) / MaxRows) + 1;

        if (targetSheet > MaxSheets)
            throw new InvalidOperationException(
                $"ماکسیمم شیت ({MaxSheets}) پر شده است. ردیف {row} قابل نوشتن نیست.");

        // شیت‌های لازم رو بساز
        while (_autoSheetSplitCurrent < targetSheet)
        {
            if (_options.ZipMode == ZipMode.StreamingZip)
            {
                var (prevStream, prevWriter) = _streamingWriters[^1];
                if (!prevWriter.IsDisposed) prevWriter.Dispose();
                prevStream.Dispose();
            }
            else // SharpCompressZip
            {
                // SheetWriter رو ببند تا buffer flush بشه
                var (prevStream, prevWriter) = _sharpWriters[^1];
                if (!prevWriter.IsDisposed) prevWriter.Dispose();
                // stream رو هم ببند تا ZipWriter بتونه entry جدید باز کنه
                prevStream?.Dispose();
            }

            _autoSheetSplitCurrent++;
            _sheets.Add(new SheetConfig
            {
                Name        = $"{_autoSheetSplitPrefix} {_autoSheetSplitCurrent}",
                FreezeRows  = _autoSheetSplitFreezeRows,
                RightToLeft = _autoSheetSplitRtl
            });
            OpenSheetWriter(_sheets.Count - 1);

            // هدر رو روی شیت جدید بنویس
            if (_autoSheetSplitHeader != null)
                for (int col = 0; col < _autoSheetSplitHeader.Length; col++)
                    GetWriter(_autoSheetSplitCurrent).Write(
                        _autoSheetSplitHeader[col], col + 1, 1);
        }

        // row رو نسبت به شیت فعلی map کن
        int rowInSheet = ((row - 1) % MaxRows) + 1;
        // اگه هدر داریم، داده از ردیف ۲ شروع میشه
        int mappedRow = _autoSheetSplitHeader != null && _autoSheetSplitCurrent > 1
            ? rowInSheet + 1
            : rowInSheet;
        return (_autoSheetSplitCurrent, mappedRow);
    }

    // ── Dispose ─────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_options.ZipMode == ZipMode.SharpCompressZip)
        {
            // اول همه شیت‌ها رو ببند
            foreach (var (stream, sw) in _sharpWriters)
            {
                if (!sw.IsDisposed) sw.Dispose();
                stream?.Dispose(); // stream رو هم ببند
            }

            // بعد static files رو بنویس
            WriteStaticEntriesSharp();

            // آخر ZipWriter رو ببند
            _zipWriter!.Dispose();
        }
        else
        {
            // همه شیت‌ها رو ببند
            foreach (var (stream, sw) in _streamingWriters)
            {
                if (!sw.IsDisposed) sw.Dispose();
                stream.Dispose();
            }

            // static files بنویس
            WriteStaticEntriesStreaming();

            _zipArchive!.Dispose();
        }

        if (!_leaveOpen)
            _outputStream.Dispose();
    }

    // ── Private ─────────────────────────────────────────────────────

    private void OpenSheetWriter(int zeroIndex)
    {
        var styleIdx = 0; // style همیشه per-cell اعمال میشه، نه default شیت

        if (_options.ZipMode == ZipMode.SharpCompressZip)
        {
            var stream = _zipWriter!.WriteToStream(
                $"xl/worksheets/sheet{zeroIndex + 1}.xml",
                new ZipWriterEntryOptions());
            var writer = new SheetWriter(stream, _sheets[zeroIndex], styleIdx, leaveStreamOpen: true);
            _sharpWriters.Add((stream, writer));
        }
        else
        {
            var entry  = _zipArchive!.CreateEntry(
                $"xl/worksheets/sheet{zeroIndex + 1}.xml",
                _options.CompressionLevel);
            var stream = entry.Open();
            var writer = new SheetWriter(stream, _sheets[zeroIndex], styleIdx, leaveStreamOpen: true);
            _streamingWriters.Add((stream, writer));
        }
    }

    private SheetWriter GetWriter(int sheetIndex /* 1-based */)
    {
        if (sheetIndex < 1 || sheetIndex > _sheets.Count)
            throw new ArgumentOutOfRangeException(nameof(sheetIndex),
                $"شیت {sheetIndex} وجود ندارد. محدوده مجاز: 1 تا {_sheets.Count}");

        return _options.ZipMode == ZipMode.SharpCompressZip
            ? _sharpWriters[sheetIndex - 1].sw
            : _streamingWriters[sheetIndex - 1].sw;
    }

    private static void ValidateCol(int col)
    {
        if (col < 1 || col > MaxCols)
            throw new ArgumentOutOfRangeException(nameof(col),
                $"ستون {col} از حد مجاز {MaxCols} بیشتر است. محدوده مجاز: 1 تا {MaxCols}.");
    }

    private static void ValidateRow(int row)
    {
        if (row < 1 || row > MaxRows)
            throw new ArgumentOutOfRangeException(nameof(row),
                $"ردیف {row} از حد مجاز خارج است. محدوده مجاز: 1 تا {MaxRows:N0}.");
    }

    private void WriteStaticEntriesStreaming()
    {
        WriteZipEntryStreaming("[Content_Types].xml",       XmlHelper.ContentTypes(_sheets.Count));
        WriteZipEntryStreaming("_rels/.rels",                XmlHelper.Rels());
        WriteZipEntryStreaming("xl/workbook.xml",            XmlHelper.Workbook(_sheets));
        WriteZipEntryStreaming("xl/_rels/workbook.xml.rels", XmlHelper.WorkbookRels(_sheets.Count));
        WriteZipEntryStreaming("xl/styles.xml",              XmlHelper.Styles(_styles));
        WriteZipEntryStreaming("xl/sharedStrings.xml",       XmlHelper.SharedStrings());
    }

    private void WriteZipEntryStreaming(string name, string content)
    {
        var entry = _zipArchive!.CreateEntry(name, _options.CompressionLevel);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(content);
    }

    private void WriteStaticEntriesSharp()
    {
        WriteZipEntrySharp("[Content_Types].xml",       XmlHelper.ContentTypes(_sheets.Count));
        WriteZipEntrySharp("_rels/.rels",                XmlHelper.Rels());
        WriteZipEntrySharp("xl/workbook.xml",            XmlHelper.Workbook(_sheets));
        WriteZipEntrySharp("xl/_rels/workbook.xml.rels", XmlHelper.WorkbookRels(_sheets.Count));
        WriteZipEntrySharp("xl/styles.xml",              XmlHelper.Styles(_styles));
        WriteZipEntrySharp("xl/sharedStrings.xml",       XmlHelper.SharedStrings());
    }

    private void WriteZipEntrySharp(string name, string content)
    {
        using var stream = _zipWriter!.WriteToStream(name, new ZipWriterEntryOptions());
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(content);
    }

    private static FileStream OpenFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("مسیر فایل نمی‌تواند خالی باشد.", nameof(filePath));
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);
        return new FileStream(filePath, FileMode.Create, FileAccess.Write);
    }
}
