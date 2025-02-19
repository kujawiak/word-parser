using System.Text.RegularExpressions;

namespace WordParser
{
    public static class StringExtensions
    {
        public static string Sanitize(this string input)
        {
            return Regex.Replace(input, @"\s+", " ");
        }

        public static string ExtractOrdinal(this string input)
        {
            var match = Regex.Match(input, @"^([^\)]+)\)");
            return match.Success ? match.Groups[1].Value : "Unknown";
        }
    }
}