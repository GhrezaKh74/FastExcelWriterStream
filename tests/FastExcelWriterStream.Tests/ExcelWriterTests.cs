using System.IO.Compression;
using FastExcelWriterStream;
using Xunit;

namespace FastExcelWriterStream.Tests;

public class ExcelWriterTests
{
    private static string TempFile() =>
        Path.Combine("", $"test_{Guid.NewGuid()}.xlsx");

    [Fact]
    public void Write_CreatesValidXlsxFile()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path))
        {
            ew.Write("Hello", 1, 1);
            ew.Write("World", 2, 1);
            ew.Write("123", 1, 2, DataType.Number);
        }

        Assert.True(File.Exists(path));
        Assert.True(new FileInfo(path).Length > 0);

        using var zip = ZipFile.OpenRead(path);
        Assert.NotNull(zip.GetEntry("xl/worksheets/sheet1.xml"));
        Assert.NotNull(zip.GetEntry("xl/workbook.xml"));
        Assert.NotNull(zip.GetEntry("xl/styles.xml"));

        File.Delete(path);
    }

    [Fact]
    public void Write_LargeDataset_DoesNotThrow()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path, new SheetConfig { Name = "Data", FreezeRows = 1 }))
        {
            for (int col = 1; col <= 13; col++)
                ew.Write($"Column{col}", col, 1);

            for (int row = 2; row <= 10_001; row++)
                for (int col = 1; col <= 13; col++)
                    ew.Write($"val_{row}_{col}", col, row);
        }

        Assert.True(new FileInfo(path).Length > 0);
        File.Delete(path);
    }

    [Fact]
    public void Write_OutOfOrderRow_ThrowsException()
    {
        var path = TempFile();
        try
        {
            using var ew = new ExcelWriter(path);
            ew.Write("A", 1, 5);
            Assert.Throws<InvalidOperationException>(() => ew.Write("B", 1, 3));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Write_WithStyle_CreatesFile()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path,
            style: new StyleConfig { Bold = true, FillColor = "4472C4", FontColor = "FFFFFF" }))
        {
            ew.Write("Header", 1, 1);
            ew.WriteNumber(42.5, 2, 1);
        }

        Assert.True(File.Exists(path));
        File.Delete(path);
    }

    [Fact]
    public void WriteToStream_Works()
    {
        using var ms = new MemoryStream();
        using (var ew = new ExcelWriter(ms, leaveOpen: true))
        {
            ew.Write("StreamTest", 1, 1);
        }

        Assert.True(ms.Length > 0);
    }

    [Fact]
    public void AddSheet_MultipleSheets_CreatesCorrectEntries()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path, new SheetConfig { Name = "Sheet1" }))
        {
            ew.Write("Sheet1 Data", 1, 1);
            var s2 = ew.AddSheet(new SheetConfig { Name = "Sheet2" });
            ew.Write("Sheet2 Data", 1, 1, s2);
        }

        using var zip = ZipFile.OpenRead(path);
        Assert.NotNull(zip.GetEntry("xl/worksheets/sheet1.xml"));
        Assert.NotNull(zip.GetEntry("xl/worksheets/sheet2.xml"));

        File.Delete(path);
    }

    [Fact]
    public void WriteFormula_BuiltIn_Works()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path))
        {
            for (int row = 1; row <= 10; row++)
                ew.WriteNumber(row, 1, row);

            ew.WriteFormula(FormulaType.Sum, 1, 12, 1, 1, 10);
            ew.WriteFormula(FormulaType.Average, 1, 13, 1, 1, 10);
        }

        Assert.True(File.Exists(path));
        File.Delete(path);
    }
    [Fact]
    public void Write_AutoSplitSheet()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path, options: new ExcelWriterOptions
        {
            ZipMode = ZipMode.StreamingZip,
            CompressionLevel = CompressionLevel.Fastest// سریع‌تر
        }))
         {
            ew.EnableAutoSheetSplit();
            for (int row = 1; row <= 3_000_000; row++)
                ew.WriteNumber(row, 1, row);
        }

        Assert.True(File.Exists(path));
        File.Delete(path);
    }
}
