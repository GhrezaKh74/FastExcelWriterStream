# Changelog

All notable changes to FastExcelWriterStream will be documented in this file.

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
