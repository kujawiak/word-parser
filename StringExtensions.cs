using System.Text.RegularExpressions;

namespace WordParser
{
    public static class StringExtensions
    {
        public static string Sanitize(this string input)
        {
            input = Regex.Replace(input, @"\s+", " ");
            return input.Replace("–", "-");
        }

        public static string ExtractOrdinal(this string input)
        {
            var match = Regex.Match(input, @"^([^\)]+)\)");
            return match.Success ? match.Groups[1].Value : "Unknown";
        }

        public static string GetAmendingProcedure(this string input)
        {
            // var match = Regex.Match(input, @"W art\. (\w+)");
            // if (match.Success)
            // {
            //     return match.Groups[1].Value;
            // }
            // else
            // {
            //     match = Regex.Match(input, @"ust\. (\d+) otrzymuje brzmienie");
            //     if (match.Success)
            //     {
            //         return match.Groups[1].Value;
            //     }
            //     else
            //     {
            //         match = Regex.Match(input, @"dodaje się (.+) w brzmieniu");
            //         return match.Success ? match.Groups[1].Value : "Unknown";
            //     }
            // }
        // Wyciąganie tytułu ustawy
        string titlePattern = @"o (.*?) \(";
        Match titleMatch = Regex.Match(input, titlePattern);
        string title = titleMatch.Groups[1].Value;

        // Wyciąganie daty wydania
        string datePattern = @"z dnia (\d{1,2} \w+ \d{4} r\.)";
        Match dateMatch = Regex.Match(input, datePattern);
        string date = dateMatch.Groups[1].Value;

        // Wyciąganie publikatorów
        string publishersPattern = @"\((Dz\. U\. z \d{4} r\. poz\. \d+ i \d+)\)";
        Match publishersMatch = Regex.Match(input, publishersPattern);
        string publishers = publishersMatch.Groups[1].Value;

        // Wyciąganie zmian
        string changesPattern = @"po art\. (\d+) dodaje się art\. (\d+\w)";
        Match changesMatch = Regex.Match(input, changesPattern);
        string originalArticle = changesMatch.Groups[1].Value;
        string newArticle = changesMatch.Groups[2].Value;

        // Wyświetlanie wyników
        Console.WriteLine($"Tytuł ustawy: {title}");
        Console.WriteLine($"Data wydania: {date}");
        Console.WriteLine($"Publikatory: {publishers}");
        Console.WriteLine($"Wprowadzone zmiany: dodanie nowego artykułu {newArticle} po artykule {originalArticle}");
        return input;
        }
    }   
}