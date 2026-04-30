# FastExcelWriter

[![NuGet](https://img.shields.io/nuget/v/FastExcelWriter.svg)](https://www.nuget.org/packages/FastExcelWriter)
[![NuGet Downloads](https://img.shields.io/nuget/dt/FastExcelWriter.svg)](https://www.nuget.org/packages/FastExcelWriter)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![.NET](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com)

Lightweight, extremely fast and memory efficient Excel (`.xlsx`) writer for .NET 8.  
Streams data **directly to file** — no XML serialization, no memory footprint.

---

## ✨ Features

- 🚀 **Extremely fast** — streams data directly to file
- 💾 **Minimal RAM** — no data accumulation in memory
- 📄 **Multiple sheets** — up to 10 sheets per file
- 🔀 **Auto sheet split** — automatically creates new sheets after 1M rows
- 🎨 **Cell styles** — bold, italic, colors, borders, alignment
- 🔢 **Number formats** — thousands separator, decimals, percent
- 📐 **Auto column width** — sample-based width detection
- 🔒 **Freeze panes** — freeze rows and columns
- ↔️ **RTL support** — right-to-left sheets
- 🔽 **AutoFilter** — column filters
- ↔️ **Merge cells** — merge cell ranges
- 📝 **Formulas** — custom and built-in (SUM, AVERAGE, etc.)

---

## 📊 Performance

| Library | Time | RAM | Rows × Cols |
|---|---|---|---|
| **FastExcelWriter** | ~7s | ~8 MB | 1M × 50 |
| EPPlus | ~44s | ~2,900 MB | 1M × 50 |
| ClosedXML | ~60s+ | ~3,500 MB | 1M × 50 |

---

## 📦 Installation

```bash
dotnet add package FastExcelWriterStream
```

---

## 🚀 Quick Start

```csharp
using FastExcelWriter;

using var ew = new ExcelWriter("output.xlsx");

ew.Write("Name",  1, 1);
ew.Write("Amount", 2, 1);

ew.Write("Alice", 1, 2);
ew.WriteNumber(22954062, 2, 2);
```

---

## 📖 Usage

### Basic Writing

```csharp
using var ew = new ExcelWriter("output.xlsx", new SheetConfig
{
    Name        = "Report",
    FreezeRows  = 1,        // freeze header row
    RightToLeft = true      // RTL sheet
});

// Write text
ew.Write("Hello", col: 1, row: 1);

// Write number
ew.WriteNumber(1234567.89, col: 2, row: 1);

// Write formula
ew.WriteFormula("A1+B1", col: 3, row: 1);
ew.WriteFormula(FormulaType.Sum, col: 4, row: 1,
                dataCol: 2, dataRowStart: 1, dataRowEnd: 100);
```

### Styles

```csharp
using var ew = new ExcelWriter("output.xlsx");

// Register styles BEFORE any Write call
int headerStyle = ew.AddStyle(new StyleConfig
{
    Bold         = true,
    FillColor    = "1F4E79",   // dark blue background
    FontColor    = "FFFFFF",   // white text
    HAlign       = HAlign.Center,
    BorderBottom = true
});

int amountStyle = ew.AddStyle(new StyleConfig
{
    NumberFormat = NumberFormat.Thousands,   // 1,234,567
    HAlign       = HAlign.Right
});

int dollarStyle = ew.AddStyle(new StyleConfig
{
    NumberFormat  = NumberFormat.Thousands,
    DecimalPlaces = 2                        // 1,234,567.00
});

// Apply styles using named parameter
ew.Write("Revenue",    1, 1, styleIndex: headerStyle);
ew.WriteNumber(22954062,  1, 2, styleIndex: amountStyle);
ew.WriteNumber(1234.56,   1, 3, styleIndex: dollarStyle);
```

### Multiple Sheets

```csharp
using var ew = new ExcelWriter("report.xlsx", new SheetConfig { Name = "Summary" });

ew.Write("Summary data", 1, 1);

// Add second sheet
int sheet2 = ew.AddSheet(new SheetConfig { Name = "Details" });
ew.Write("Detail data", 1, 1, sheet2);
```

### Auto Sheet Split (3M+ rows)

```csharp
using var ew = new ExcelWriter("big_report.xlsx");

string[] headers = { "ID", "Name", "Amount", "Date" };

ew.EnableAutoSheetSplit(
    sheetNamePrefix: "Data",         // → "Data 1", "Data 2", "Data 3"
    headerRow:       headers,        // auto-repeat on each new sheet
    freezeRows:      1,
    rightToLeft:     true
);

// Write header on first sheet manually
for (int col = 1; col <= headers.Length; col++)
    ew.Write(headers[col - 1], col, 1);

// Write 3 million rows — sheet splitting is automatic
for (int row = 2; row <= 3_000_001; row++)
{
    ew.WriteNumber(row - 1,       1, row);
    ew.Write($"Name {row}",       2, row);
    ew.WriteNumber(row * 1000.0,  3, row);
    ew.Write("2026/01/01",        4, row);
}
```

### Auto Column Width

```csharp
using var ew = new ExcelWriter("output.xlsx");

// Must be called BEFORE any Write
ew.EnableAutoWidth(maxSampleRows: 1000);  // sample first 1000 rows

ew.Write("Transaction Description", 1, 1);
// ... rest of writes
```

### Write to Stream (Azure Blob, etc.)

```csharp
using var stream = await blobClient.OpenWriteAsync(overwrite: true);
using var ew = new ExcelWriter(stream, leaveOpen: false);

ew.Write("Hello", 1, 1);
```

### SheetConfig Options

```csharp
new SheetConfig
{
    Name         = "Sheet1",
    FreezeRows   = 1,                          // freeze top N rows
    FreezeCols   = 0,                          // freeze left N columns
    RightToLeft  = false,                      // RTL direction
    ColumnsWidth = new List<double>            // manual column widths
    {
        10, 25, 15, 12
    },
    MergeRanges  = new List<MergeRange>        // merge cells
    {
        new MergeRange(colStart: 1, rowStart: 1, colEnd: 4, rowEnd: 1)
    },
    FilterRange  = new FilterRange(            // AutoFilter
        colStart: 1, rowStart: 1,
        colEnd:   13, rowEnd:  1
    )
}
```

### ExcelWriterOptions

```csharp
new ExcelWriterOptions
{
    // StreamingZip: low RAM, streaming (default — recommended for large files)
    // SharpCompressZip: faster, higher RAM
    ZipMode = ZipMode.StreamingZip,

    // Optimal (default) / Fastest / NoCompression
    CompressionLevel = CompressionLevel.Fastest
}
```

---

## 📋 API Reference

### ExcelWriter

| Method | Description |
|---|---|
| `Write(value, col, row, dataType?, styleIndex?)` | Write text/number/formula to sheet 1 |
| `WriteNumber(value, col, row, styleIndex?)` | Write numeric value to sheet 1 |
| `WriteFormula(formula, col, row)` | Write formula string to sheet 1 |
| `WriteFormula(type, col, row, dataCol, start, end)` | Write built-in formula |
| `Write(value, col, row, sheetIndex, ...)` | Write to specific sheet |
| `WriteNumber(value, col, row, sheetIndex, ...)` | Write number to specific sheet |
| `AddSheet(SheetConfig)` | Add new sheet, returns 1-based index |
| `AddStyle(StyleConfig)` | Register style, returns style index |
| `EnableAutoSheetSplit(...)` | Enable auto sheet splitting for 1M+ rows |
| `EnableAutoWidth(maxSampleRows?)` | Enable automatic column width |
| `BeginAutoSplit(cols, prefix?)` | Begin column-auto-tracked writing |
| `AutoWrite(value, dataType?)` | Write next cell (auto col/row/sheet) |
| `AutoWriteNumber(value)` | Write next numeric cell |

### StyleConfig Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `Bold` | bool | false | Bold text |
| `Italic` | bool | false | Italic text |
| `Underline` | bool | false | Underlined text |
| `FontSize` | int | 11 | Font size in points |
| `FontName` | string | "Calibri" | Font name |
| `FontColor` | string? | null | Hex color e.g. `"FF0000"` |
| `FillColor` | string? | null | Background hex e.g. `"4472C4"` |
| `BorderLeft/Right/Top/Bottom` | bool | false | Border sides |
| `BorderColor` | string | "000000" | Border hex color |
| `HAlign` | HAlign | General | Horizontal alignment |
| `VAlign` | VAlign | Bottom | Vertical alignment |
| `NumberFormat` | NumberFormat | None | Number display format |
| `DecimalPlaces` | int | 0 | Decimal places (0–10) |

### Enumerations

```csharp
enum DataType    { Text, Number, Formula }
enum FormulaType { Average, Count, Max, Min, Sum }
enum HAlign      { General, Left, Center, Right, Fill, Justify }
enum VAlign      { Top, Center, Bottom, Justify, Distributed }
enum NumberFormat { None, Thousands, ThousandsDecimal, Percent, Custom }
enum ZipMode     { StreamingZip, SharpCompressZip }
```

---

## ⚠️ Important Notes

1. **Ascending row order** — rows must be written from low to high
2. **AddStyle before Write** — register all styles before any Write call
3. **Named parameter** — always use `styleIndex: myStyle`, not positional
4. **EnableAutoWidth before Write** — must be called before any Write
5. **using statement** — always use `using` to ensure file is properly closed
6. **StreamingZip + AddSheet** — previous sheet must be fully written before adding new sheet

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.
