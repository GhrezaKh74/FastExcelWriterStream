using System.Text;

namespace FastExcelWriter;

internal static class StylesBuilder
{
    internal static string Build(StyleConfig s)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        // fonts
        sb.Append("<fonts count=\"2\">");
        sb.Append("<font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
        sb.Append("<font>");
        if (s.Bold)      sb.Append("<b/>");
        if (s.Italic)    sb.Append("<i/>");
        if (s.Underline) sb.Append("<u/>");
        sb.Append($"<sz val=\"{s.FontSize}\"/>");
        if (!string.IsNullOrEmpty(s.FontColor))
            sb.Append($"<color rgb=\"FF{s.FontColor!.TrimStart('#')}\"/>");
        sb.Append($"<name val=\"{s.FontName}\"/>");
        sb.Append("</font>");
        sb.Append("</fonts>");

        // fills
        sb.Append("<fills count=\"3\">");
        sb.Append("<fill><patternFill patternType=\"none\"/></fill>");
        sb.Append("<fill><patternFill patternType=\"gray125\"/></fill>");
        sb.Append(!string.IsNullOrEmpty(s.FillColor)
            ? $"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{s.FillColor!.TrimStart('#')}\"/></patternFill></fill>"
            : "<fill><patternFill patternType=\"none\"/></fill>");
        sb.Append("</fills>");

        // borders
        var bc = $"FF{s.BorderColor.TrimStart('#')}";
        sb.Append("<borders count=\"2\">");
        sb.Append("<border><left/><right/><top/><bottom/></border>");
        sb.Append("<border>");
        sb.Append(s.BorderLeft   ? $"<left style=\"thin\"><color rgb=\"{bc}\"/></left>"    : "<left/>");
        sb.Append(s.BorderRight  ? $"<right style=\"thin\"><color rgb=\"{bc}\"/></right>"  : "<right/>");
        sb.Append(s.BorderTop    ? $"<top style=\"thin\"><color rgb=\"{bc}\"/></top>"      : "<top/>");
        sb.Append(s.BorderBottom ? $"<bottom style=\"thin\"><color rgb=\"{bc}\"/></bottom>": "<bottom/>");
        sb.Append("</border>");
        sb.Append("</borders>");

        // cellStyleXfs
        sb.Append("<cellStyleXfs count=\"1\">");
        sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>");
        sb.Append("</cellStyleXfs>");

        // cellXfs
        var hasAlign = s.HAlign != HAlign.General || s.VAlign != VAlign.Bottom;
        var alignXml = hasAlign
            ? $"<alignment horizontal=\"{s.HAlign.ToString().ToLowerInvariant()}\" vertical=\"{s.VAlign.ToString().ToLowerInvariant()}\"/>"
            : string.Empty;

        sb.Append("<cellXfs>");
        sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
        sb.Append(string.IsNullOrEmpty(alignXml)
            ? "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\"/>"
            : $"<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\">{alignXml}</xf>");
        sb.Append("</cellXfs>");

        sb.Append("</styleSheet>");
        return sb.ToString();
    }
}
