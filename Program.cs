using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser
{
    class Program
    {
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
                    var legalAct = new WordParserLibrary.LegalAct(wordDoc);

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
                    }
                    else if (option == "--generatexml")
                    {
                        legalAct.GenerateXML();
                        legalAct.SaveAmendmentList();
                        legalAct.CommentErrors();
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
                // System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                // {
                //     FileName = filePath,
                //     UseShellExecute = true
                // });
            }
            Console.ReadLine(); 
        }
    }
}