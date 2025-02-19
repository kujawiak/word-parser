using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser
{
    public static class ParagraphExtensions
    {
        public static string? StyleId(this Paragraph paragraph)
        {
            return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.ToString();
        }
        public static bool? StyleId(this Paragraph paragraph, string styleId)
        {
            return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.ToString()?.StartsWith(styleId);
        }
    }
}