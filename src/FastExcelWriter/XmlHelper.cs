using System.Linq;
using System.Text;

namespace FastExcelWriter;

internal static class XmlHelper
{
    internal static string EscapeXml(string s) =>
        s.Replace("&", "&amp;")
         .Replace("<", "&lt;")
         .Replace(">", "&gt;")
         .Replace("\"", "&quot;")
         .Replace("'", "&apos;");

    internal static string ColName(int col)
    {
        var result = "";
        while (col > 0) { col--; result = (char)('A' + col % 26) + result; col /= 26; }
        return result;
    }

    internal static string ContentTypes(int sheetCount)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        sb.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
        for (int i = 1; i <= sheetCount; i++)
            sb.Append($"<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        sb.Append("</Types>");
        return sb.ToString();
    }

    internal static string Rels() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
        "</Relationships>";

    internal static string Workbook(IList<SheetConfig> sheets)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        sb.Append("<sheets>");
        for (int i = 0; i < sheets.Count; i++)
            sb.Append($"<sheet name=\"{EscapeXml(sheets[i].Name)}\" sheetId=\"{i+1}\" r:id=\"rId{i+1}\"/>");
        sb.Append("</sheets></workbook>");
        return sb.ToString();
    }

    internal static string WorkbookRels(int sheetCount)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (int i = 1; i <= sheetCount; i++)
            sb.Append($"<Relationship Id=\"rId{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{sheetCount+1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{sheetCount+2}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    internal static string SharedStrings() =>
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></sst>";

    internal static string Styles(IList<StyleConfig> styles)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        // ── numFmts باید اول بیاد ──────────────────────────────────
        var customFmts = styles.Where(s => s.NumberFormat != NumberFormat.None || s.DecimalPlaces > 0).ToList();
        if (customFmts.Count > 0)
        {
            var fmtCount = 0;
            if (styles.Any(s => s.NumberFormat == NumberFormat.Thousands && s.DecimalPlaces == 0)) fmtCount++;
            if (styles.Any(s => s.NumberFormat == NumberFormat.ThousandsDecimal && s.DecimalPlaces == 0)) fmtCount++;
            if (styles.Any(s => s.NumberFormat == NumberFormat.Percent)) fmtCount++;
            var customDecimals = styles.Where(s => s.DecimalPlaces > 0).Select(s => s.DecimalPlaces).Distinct().OrderBy(x => x).ToList();
            fmtCount += customDecimals.Count;

            sb.Append($"<numFmts count=\"{fmtCount}\">");
            if (styles.Any(s => s.NumberFormat == NumberFormat.Thousands && s.DecimalPlaces == 0))
                sb.Append("<numFmt numFmtId=\"164\" formatCode=\"#,##0\"/>");
            if (styles.Any(s => s.NumberFormat == NumberFormat.ThousandsDecimal && s.DecimalPlaces == 0))
                sb.Append("<numFmt numFmtId=\"165\" formatCode=\"#,##0.00\"/>");
            if (styles.Any(s => s.NumberFormat == NumberFormat.Percent))
                sb.Append("<numFmt numFmtId=\"166\" formatCode=\"0.00%\"/>");
            int customId = 200;
            foreach (var dec in customDecimals)
            {
                var zeros = new string('0', dec);
                sb.Append($"<numFmt numFmtId=\"{customId}\" formatCode=\"#,##0.{zeros}\"/>");
                customId++;
            }
            sb.Append("</numFmts>");
        }

        // ── fonts: index 0=default, 1..n=custom ───────────────────
        sb.Append($"<fonts count=\"{styles.Count + 1}\">");
        sb.Append("<font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
        foreach (var s in styles)
        {
            sb.Append("<font>");
            if (s.Bold)      sb.Append("<b/>");
            if (s.Italic)    sb.Append("<i/>");
            if (s.Underline) sb.Append("<u/>");
            sb.Append($"<sz val=\"{s.FontSize}\"/>");
            if (!string.IsNullOrEmpty(s.FontColor))
                sb.Append($"<color rgb=\"FF{s.FontColor!.TrimStart('#')}\"/>");
            sb.Append($"<name val=\"{s.FontName}\"/>");
            sb.Append("</font>");
        }
        sb.Append("</fonts>");

        // ── fills: 0=none, 1=gray125, 2..n=custom ─────────────────
        sb.Append($"<fills count=\"{styles.Count + 2}\">");
        sb.Append("<fill><patternFill patternType=\"none\"/></fill>");
        sb.Append("<fill><patternFill patternType=\"gray125\"/></fill>");
        foreach (var s in styles)
        {
            sb.Append(!string.IsNullOrEmpty(s.FillColor)
                ? $"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{s.FillColor!.TrimStart('#')}\"/></patternFill></fill>"
                : "<fill><patternFill patternType=\"none\"/></fill>");
        }
        sb.Append("</fills>");

        // ── borders: 0=default, 1..n=custom ───────────────────────
        sb.Append($"<borders count=\"{styles.Count + 1}\">");
        sb.Append("<border><left/><right/><top/><bottom/></border>");
        foreach (var s in styles)
        {
            var bc = $"FF{s.BorderColor.TrimStart('#')}";
            sb.Append("<border>");
            sb.Append(s.BorderLeft   ? $"<left style=\"thin\"><color rgb=\"{bc}\"/></left>"    : "<left/>");
            sb.Append(s.BorderRight  ? $"<right style=\"thin\"><color rgb=\"{bc}\"/></right>"  : "<right/>");
            sb.Append(s.BorderTop    ? $"<top style=\"thin\"><color rgb=\"{bc}\"/></top>"      : "<top/>");
            sb.Append(s.BorderBottom ? $"<bottom style=\"thin\"><color rgb=\"{bc}\"/></bottom>": "<bottom/>");
            sb.Append("</border>");
        }
        sb.Append("</borders>");

        // ── cellStyleXfs ───────────────────────────────────────────
        sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

        // ── cellXfs: xf[0]=default, xf[1..n]=custom ───────────────
        sb.Append("<cellXfs>");
        sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
        var customDecimalsAll = styles.Where(x => x.DecimalPlaces > 0).Select(x => x.DecimalPlaces).Distinct().OrderBy(x => x).ToList();
        for (int i = 0; i < styles.Count; i++)
        {
            var s = styles[i];
            int numFmtId;
            if (s.DecimalPlaces > 0)
                numFmtId = 200 + customDecimalsAll.IndexOf(s.DecimalPlaces);
            else
                numFmtId = (int)s.NumberFormat;

            var alignPart = (s.HAlign != HAlign.General || s.VAlign != VAlign.Bottom)
                ? $"<alignment horizontal=\"{s.HAlign.ToString().ToLowerInvariant()}\" vertical=\"{s.VAlign.ToString().ToLowerInvariant()}\"/>"
                : "";
            sb.Append(string.IsNullOrEmpty(alignPart)
                ? $"<xf numFmtId=\"{numFmtId}\" fontId=\"{i+1}\" fillId=\"{i+2}\" borderId=\"{i+1}\" xfId=\"0\"/>"
                : $"<xf numFmtId=\"{numFmtId}\" fontId=\"{i+1}\" fillId=\"{i+2}\" borderId=\"{i+1}\" xfId=\"0\">{alignPart}</xf>");
        }
        sb.Append("</cellXfs>");
        sb.Append("</styleSheet>");
        return sb.ToString();
    }
}
