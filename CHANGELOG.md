# Changelog

All notable changes to FastExcelWriterStream will be documented in this file.

## [2.2.0] - 2026-06-18

### Added
- `WriteDate(DateTime, ...)` — dates stored as Excel serial numbers with a built-in date format
- `WriteBool(bool, ...)` — native boolean cells, plus `DataType.Boolean`
- Built-in date/time number formats: `NumberFormat.Date`, `NumberFormat.Time`, `NumberFormat.DateTime`
- High-level `WriteRow(...)` / `WriteRowValues(...)` with automatic type detection
  (string, number, `DateTime`/`DateOnly`/`TimeOnly`, bool); `null` leaves a cell empty
- `WriteRecords<T>(...)` to export an `IEnumerable<T>` as a table, with
  `[ExcelColumn]` (name/order/format) and `[ExcelIgnore]` attributes
- `CurrentRow` property exposing the next high-level row

### Changed
- Raised `MaxCols` to 16384 and `MaxSheets` to 1000
- Updated `SharpCompress` 0.47.3 → 0.49.1 (resolves GHSA-6c8g-7p36-r338)

### Fixed
- **Illegal XML control characters** (e.g. `0x01`) in cell text are now stripped
  instead of being written verbatim — previously such characters (common in raw
  log data) produced a corrupt `.xlsx` that Excel refused to open
- Styles builder no longer emits an empty `<numFmts>` block when only
  built-in formats (date/time) are used

## [2.1.0] - 2026-06-18

### Changed
- Upgraded target framework from .NET 8 to **.NET 10** (latest LTS)
- Updated CI publish workflow to use the .NET 10 SDK
- Updated test dependencies: `Microsoft.NET.Test.Sdk` 17.12.0, `xunit` 2.9.2, `xunit.runner.visualstudio` 2.8.2

## [1.0.0] - 2026-04-30

### Added
- Streaming `.xlsx` writer for .NET 8
- Multiple sheets support (up to 10)
- Auto sheet split for datasets exceeding 1,048,576 rows
- Cell styles: bold, italic, underline, font color, background color, borders
- Number formats: thousands separator, decimal places, percent
- Auto column width with configurable sample rows
- Freeze panes (rows and columns)
- Right-to-left (RTL) sheet direction
- AutoFilter support
- Merge cells
- Built-in formulas (SUM, AVERAGE, COUNT, MAX, MIN)
- Custom formula strings
- Two ZIP modes: StreamingZip (low RAM) and SharpCompressZip (faster)
- Write to file path or any Stream (Azure Blob, MemoryStream, etc.)
