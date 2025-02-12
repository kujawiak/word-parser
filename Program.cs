using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser
{
    class Program
    {
        public static List<string> VALID_STYLES_BOLD = new List<string>
        {
            "_IG_ - indeks górny", 
            "_IG_K_ - indeks górny i kursywa", 
            "_IG_P_ - indeks górny i pogrubienie", 
            "_IG_P_K_ - indeks górny i pogrubienie kursywa", 
            "_IIG_ - indeks górny indeksu górnego", 
            "_IIG_P_ - indeks górny indeksu górnego i pogrubienie"
        };
        public static List<string> VALID_STYLES_SUPERSCRIPT = new List<string>
        {
            "_P_ - pogrubienie", 
            "_IG_P_ - indeks górny i pogrubienie", 
            "TYT(DZ)_PRZEDM - przedmiot regulacji tytułu lub działu", 
            "_IG_P_K_ - indeks górny i pogrubienie kursywa", 
            "OZN_RODZ_AKTU - tzn. ustawa lub rozporządzenie i organ wydający", 
            "_IIG_P_ - indeks górny indeksu górnego i pogrubienie", 
            "TYTUŁ_AKTU - przedmiot regulacji ustawy lub rozporządzenia", 
            "ROZDZ(ODDZ)_PRZEDM - przedmiot regulacji rozdziału lub oddziału", 
            "TYT_TABELI - tytuł tabeli", 
            "NAZ_ORG_WYD - nazwa organu wydającego projektowany akt", 
            "NAZ_ORG_W_POROZUMIENIU - nazwa organu w porozumieniu z którym akt jest wydawany", 
            "OZN_ZAŁĄCZNIKA - wskazanie nr załącznika"
        };

        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Użycie: WordParser --hyperlinks|--formatting <nazwa-pliku>");
                return;
            }

            string option = args[0];
            string filePath = args[1];

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Plik nie istnieje.");
                return;
            } else
            {
                string directoryName = Path.GetDirectoryName(filePath) ?? string.Empty;
                if (directoryName == null)
                {
                    Console.WriteLine("Nie można uzyskać katalogu z podanej ścieżki pliku.");
                    return;
                }
                string backupDirectory = directoryName;

                string backupFileName = Path.GetFileNameWithoutExtension(filePath) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(filePath);
                string backupFilePath = Path.Combine(backupDirectory, backupFileName);

                File.Copy(filePath, backupFilePath);
                filePath = backupFilePath;
            
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    wordDoc.CompressionOption = CompressionOption.Maximum;
                    var legalAct = new LegalAct(wordDoc);

                    // Usuwanie komentarzy autora 'System'
                    legalAct.RemoveSystemComments();

                    if (option == "--hyperlinks")
                    {
                        int commentCount = legalAct.ParseHyperlinks();
                        Console.WriteLine($"Liczba dodanych komentarzy: {commentCount}");
                    }
                    else if (option == "--formatting")
                    {
                        legalAct.CleanParagraphProperties();
                        legalAct.MergeRuns();
                        legalAct.MergeTexts();
                        //
                        legalAct.GenerateXMLSchema();
                    }
                    else
                    {
                        Console.WriteLine("Nieznany przełącznik. Użycie: WordParser --hyperlinks|--formatting <nazwa-pliku>");
                    }
                    string newFileName = Path.GetFileName(filePath);
                    
                    string copiesDirectory = Path.Combine(directoryName, "kopie");
                    if (!Directory.Exists(copiesDirectory))
                    {
                        Directory.CreateDirectory(copiesDirectory);
                    }
                    string newFilePath = Path.Combine(copiesDirectory, newFileName);

                    // Zapisz dokument pod nową nazwą
                    legalAct.SaveAs(newFilePath);
                }
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            Console.ReadLine(); 
        }
    }
}