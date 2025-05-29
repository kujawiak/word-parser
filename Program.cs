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
                    legalAct.CommentManager.RemoveSystemComments();

                    if (option == "--hyperlinks")
                    {
                        int commentCount = legalAct.DocumentProcessor.ParseHyperlinks();
                        Console.WriteLine($"Liczba dodanych komentarzy: {commentCount}");
                    }
                    else if (option == "--formatting")
                    {
                        legalAct.DocumentProcessor.CleanParagraphProperties();
                        legalAct.DocumentProcessor.MergeRuns();
                        legalAct.DocumentProcessor.MergeTexts();
                    }
                    else if (option == "--generatexml")
                    {
                        legalAct.XmlGenerator.Generate();
                        legalAct.SaveAmendmentList();
                        legalAct.CommentManager.CommentErrors(legalAct);
                    }
                    else if (option == "--docx")
                    {
                        legalAct.DocxGenerator.Generate();
                    }
                    else if (option == "--createAmendmentsTable")
                    {
                        using var stream = legalAct.XlsxGenerator.GenerateXlsx();
                        string xlsxFileName = Path.GetFileNameWithoutExtension(filePath) + "_amendments.xlsx";
                        string xlsxFilePath = Path.Combine(directoryName, xlsxFileName);

                        using (var fileStream = File.Create(xlsxFilePath))
                        {
                            stream.CopyTo(fileStream);
                        }

                        Console.WriteLine($"Utworzono plik: {xlsxFilePath}");
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