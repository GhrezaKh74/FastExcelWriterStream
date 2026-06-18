using System.IO.Compression;
using System.Text;
using FastExcelWriterStream;
using Xunit;

namespace FastExcelWriterStream.Tests;

public class NewFeaturesTests
{
    private static string TempFile() =>
        Path.Combine("", $"feat_{Guid.NewGuid()}.xlsx");

    private static string ReadEntry(string path, string entry)
    {
        using var zip = ZipFile.OpenRead(path);
        using var s = zip.GetEntry(entry)!.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
    }

    [Fact]
    public void WriteBool_ProducesBooleanCell()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path))
            ew.WriteBool(true, 1, 1);

        var xml = ReadEntry(path, "xl/worksheets/sheet1.xml");
        Assert.Contains("t=\"b\"", xml);
        Assert.Contains("<v>1</v>", xml);
        File.Delete(path);
    }

    [Fact]
    public void WriteDate_StoresSerialNumberAndDateFormat()
    {
        var path = TempFile();
        var date = new DateTime(2024, 1, 15);
        using (var ew = new ExcelWriter(path))
            ew.WriteDate(date, 1, 1);

        var sheet  = ReadEntry(path, "xl/worksheets/sheet1.xml");
        var styles = ReadEntry(path, "xl/styles.xml");

        // serial number for 2024-01-15
        Assert.Contains($"<v>{date.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture)}</v>", sheet);
        // built-in date format id 14
        Assert.Contains("numFmtId=\"14\"", styles);
        File.Delete(path);
    }

    [Fact]
    public void WriteRow_AutoDetectsTypes()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path))
        {
            int written = ew.WriteRow("Name", 42, true, new DateTime(2024, 5, 1));
            Assert.Equal(1, written);
            Assert.Equal(2, ew.CurrentRow);
            ew.WriteRow("Second", 3.14);
        }

        var xml = ReadEntry(path, "xl/worksheets/sheet1.xml");
        Assert.Contains("inlineStr", xml);   // text
        Assert.Contains("<v>42</v>", xml);    // number
        Assert.Contains("t=\"b\"", xml);      // bool
        File.Delete(path);
    }

    [Fact]
    public void WriteRow_NullLeavesCellEmpty()
    {
        var path = TempFile();
        using (var ew = new ExcelWriter(path))
            ew.WriteRow("A", null, "C");

        var xml = ReadEntry(path, "xl/worksheets/sheet1.xml");
        Assert.Contains("A1", xml);
        Assert.Contains("C1", xml);
        Assert.DoesNotContain("B1", xml);    // null skipped
        File.Delete(path);
    }

    private sealed class Person
    {
        [ExcelColumn("Full Name", Order = 1)]
        public string Name { get; set; } = "";

        [ExcelColumn("Salary", Order = 2, Format = NumberFormat.Thousands)]
        public int Salary { get; set; }

        [ExcelColumn("Joined", Order = 3)]
        public DateTime JoinDate { get; set; }

        [ExcelIgnore]
        public string Secret { get; set; } = "hidden";
    }

    [Fact]
    public void WriteRecords_WritesHeaderAndRowsAndHonorsAttributes()
    {
        var path = TempFile();
        var people = new[]
        {
            new Person { Name = "Ali",  Salary = 1500000, JoinDate = new DateTime(2022, 3, 1) },
            new Person { Name = "Sara", Salary = 2300000, JoinDate = new DateTime(2023, 9, 15) },
        };

        int count;
        using (var ew = new ExcelWriter(path))
            count = ew.WriteRecords(people);

        Assert.Equal(2, count);

        var xml = ReadEntry(path, "xl/worksheets/sheet1.xml");
        Assert.Contains("Full Name", xml);   // custom header
        Assert.Contains("Salary", xml);
        Assert.Contains("Joined", xml);
        Assert.DoesNotContain("hidden", xml); // [ExcelIgnore]
        Assert.Contains("Ali", xml);
        Assert.Contains("Sara", xml);
        File.Delete(path);
    }

    [Fact]
    public void Write_StripsIllegalControlChars_ProducesValidXml()
    {
        var path = TempFile();
        // متن لاگ با کاراکتر کنترلی غیرمجاز 0x01 (که اکسل رو می‌شکنه)
        var dirty = "\u0001{\"Title\":\"log\"}\u0001";
        using (var ew = new ExcelWriter(path))
        {
            ew.Write(dirty, 1, 1);
            ew.WriteRow(dirty, "ok");
        }

        var xml = ReadEntry(path, "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain('\u0001', xml);   // کاراکتر غیرمجاز پاک شده
        Assert.Contains("{&quot;Title&quot;:&quot;log&quot;}", xml);

        // باید XML معتبر باشه (وگرنه اکسل باز نمی‌کنه)
        var doc = new System.Xml.XmlDocument();
        using (var zip = ZipFile.OpenRead(path))
        using (var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
            doc.Load(s);   // اگه نامعتبر بود، اینجا throw می‌کنه

        File.Delete(path);
    }

    [Fact]
    public void WriteRecords_EmptyType_Throws()
    {
        var path = TempFile();
        try
        {
            using var ew = new ExcelWriter(path);
            Assert.Throws<InvalidOperationException>(() =>
                ew.WriteRecords(new[] { new object() }));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
