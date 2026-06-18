using DotNetCompressionLevel = System.IO.Compression.CompressionLevel;

namespace FastExcelWriterStream;

/// <summary>Sheet configuration</summary>
public sealed class SheetConfig
{
    public string Name { get; set; } = "Sheet1";
    public List<double>? ColumnsWidth { get; set; }
    public List<double?>? RowsHeight { get; set; }
    public int  FreezeRows   { get; set; } = 0;
    public int  FreezeCols   { get; set; } = 0;
    public bool         RightToLeft  { get; set; } = false;
    public List<MergeRange>? MergeRanges { get; set; }
    public FilterRange? FilterRange { get; set; }
}

/// <summary>Cell style configuration</summary>
public sealed class StyleConfig
{
    public bool   Bold        { get; set; }
    public bool   Italic      { get; set; }
    public bool   Underline   { get; set; }
    public int    FontSize    { get; set; } = 11;
    public string FontName    { get; set; } = "Calibri";
    public string? FontColor  { get; set; }
    public string? FillColor  { get; set; }
    public bool   BorderLeft   { get; set; }
    public bool   BorderRight  { get; set; }
    public bool   BorderTop    { get; set; }
    public bool   BorderBottom { get; set; }
    public string BorderColor { get; set; } = "000000";
    public HAlign       HAlign        { get; set; } = HAlign.General;
    public VAlign       VAlign        { get; set; } = VAlign.Bottom;

    /// <summary>
    /// فرمت عددی — برای سه‌رقم سه‌رقم از Thousands استفاده کن
    /// اگه None باشه و DecimalPlaces > 0 باشه، خودکار فرمت ساخته میشه
    /// </summary>
    public NumberFormat NumberFormat  { get; set; } = NumberFormat.None;

    /// <summary>
    /// تعداد رقم اعشار (0 تا 10)
    /// مثال: DecimalPlaces = 3 → #,##0.000
    /// اگه NumberFormat.None باشه و این مقدار > 0 باشه، فرمت #,##0.000... ساخته میشه
    /// </summary>
    public int          DecimalPlaces { get; set; } = 0;

    internal bool HasStyle =>
        Bold || Italic || Underline || FontSize != 11 || FontName != "Calibri" ||
        !string.IsNullOrEmpty(FontColor) || !string.IsNullOrEmpty(FillColor) ||
        BorderLeft || BorderRight || BorderTop || BorderBottom ||
        HAlign != HAlign.General || VAlign != VAlign.Bottom;
}

/// <summary>Defines a cell merge range</summary>
public sealed class MergeRange
{
    public int ColStart { get; }
    public int RowStart { get; }
    public int ColEnd   { get; }
    public int RowEnd   { get; }
    public MergeRange(int colStart, int rowStart, int colEnd, int rowEnd)
    { ColStart=colStart; RowStart=rowStart; ColEnd=colEnd; RowEnd=rowEnd; }
}

/// <summary>Defines the AutoFilter range</summary>
public sealed class FilterRange
{
    public int ColStart { get; }
    public int RowStart { get; }
    public int ColEnd   { get; }
    public int RowEnd   { get; }
    public FilterRange(int colStart, int rowStart, int colEnd, int rowEnd)
    { ColStart=colStart; RowStart=rowStart; ColEnd=colEnd; RowEnd=rowEnd; }
}

/// <summary>فرمت عددی آماده برای نمایش اعداد</summary>
public enum NumberFormat
{
    /// <summary>بدون فرمت خاص (پیش‌فرض)</summary>
    None             = 0,
    /// <summary>تاریخ کوتاه (built-in اکسل): mm-dd-yy</summary>
    Date             = 14,
    /// <summary>زمان (built-in اکسل): h:mm:ss</summary>
    Time             = 21,
    /// <summary>تاریخ و زمان (built-in اکسل): mm-dd-yy h:mm</summary>
    DateTime         = 22,
    /// <summary>عدد صحیح با جداکننده سه‌رقم: 1,234,567</summary>
    Thousands        = 164,
    /// <summary>عدد اعشاری با جداکننده سه‌رقم: 1,234,567.00</summary>
    ThousandsDecimal = 165,
    /// <summary>درصد: 12.50%</summary>
    Percent          = 166,
    /// <summary>فرمت سفارشی — تعداد اعشار از DecimalPlaces خونده میشه</summary>
    Custom           = 167,
}

/// <summary>Cell data type</summary>
public enum DataType { Text, Number, Formula, Boolean }

/// <summary>
/// نگاشت یک property به ستون اکسل برای WriteRecords.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelColumnAttribute : Attribute
{
    /// <summary>عنوان ستون (هدر). اگه خالی باشه از نام property استفاده میشه.</summary>
    public string? Name { get; set; }

    /// <summary>ترتیب ستون. ستون‌های با Order کمتر اول میان. پیش‌فرض = ترتیب تعریف.</summary>
    public int Order { get; set; } = int.MaxValue;

    /// <summary>فرمت عددی این ستون.</summary>
    public NumberFormat Format { get; set; } = NumberFormat.None;

    /// <summary>تعداد رقم اعشار این ستون.</summary>
    public int DecimalPlaces { get; set; } = 0;

    public ExcelColumnAttribute() { }
    public ExcelColumnAttribute(string name) { Name = name; }
}

/// <summary>property علامت‌گذاری‌شده در خروجی WriteRecords نادیده گرفته میشه.</summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelIgnoreAttribute : Attribute { }

/// <summary>Built-in formula types</summary>
public enum FormulaType { Average, Count, Max, Min, Sum }

/// <summary>Horizontal text alignment</summary>
public enum HAlign { General, Left, Center, Right, Fill, Justify }

/// <summary>Vertical text alignment</summary>
public enum VAlign { Top, Center, Bottom, Justify, Distributed }

/// <summary>
/// تنظیمات ExcelWriter هنگام ساخت
/// </summary>
public sealed class ExcelWriterOptions
{
    /// <summary>
    /// حالت ZIP:
    /// - StreamingZip (پیش‌فرض): ZipArchive استاندارد .NET — streaming واقعی، RAM کم، سرعت متوسط
    /// - SharpCompressZip: ZipWriter از SharpCompress — سرعت بیشتر، ولی RAM بیشتر (شیت توی MemoryStream)
    /// </summary>
    public ZipMode ZipMode { get; set; } = ZipMode.StreamingZip;

    /// <summary>
    /// سطح compression:
    /// - Optimal: فشرده‌سازی بیشتر، فایل کوچکتر، کمی کندتر
    /// - Fastest: فشرده‌سازی کمتر، فایل بزرگتر، سریع‌تر
    /// - NoCompression: بدون فشرده‌سازی، فایل بزرگ، سریع‌ترین
    /// </summary>
    public DotNetCompressionLevel CompressionLevel { get; set; } = DotNetCompressionLevel.Optimal;
}

/// <summary>حالت ZIP</summary>
public enum ZipMode
{
    /// <summary>ZipArchive استاندارد .NET — streaming واقعی، RAM کم</summary>
    StreamingZip,

    /// <summary>ZipWriter از SharpCompress — سرعت بیشتر، RAM بیشتر</summary>
    SharpCompressZip
}
