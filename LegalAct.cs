using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using WordParser.Model;

namespace WordParser
{
    public class LegalAct
    {
        public WordprocessingDocument _wordDoc
         { get; }
        public MainDocumentPart MainPart { get; }
        public DocumentSettingsPart? SettingsPart { get; }

        public Title Title { get; set; }
        public List<Article> Articles { get; set; } = new List<Article>();

        public LegalAct(WordprocessingDocument wordDoc)
        {
            _wordDoc = wordDoc;
            MainPart = _wordDoc.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is null.");
            Title = new Title(MainPart.Document.Descendants<Paragraph>()
                                        .Where(p => p.ParagraphProperties != null 
                                                && p.StyleId("TYTUAKT") == true).FirstOrDefault() ?? throw new InvalidOperationException("Title paragraph not found"));                                        // Tytuł #1
            // Title.Parts.Add(new Part());                                // Dział #1
            // Title.Parts[0].Chapters.Add(new Chapter());                 // Rozdział #1
            // Title.Parts[0].Chapters[0].Sections.Add(new Section());     // Oddział #1
            
            foreach (var paragraph in MainPart.Document.Descendants<Paragraph>()
                                                        .Where(p => p.InnerText.StartsWith("Art."))
                                                        .ToList())
            {
                if (paragraph.ParagraphProperties == null)
                {
                    Console.WriteLine("[CTOR]\tBrak właściwości paragrafu!");
                    continue;
                }
                if (paragraph.ParagraphProperties.ParagraphStyleId == null)
                {
                    Console.WriteLine("[CTOR]\tBrak stylu paragrafu!");
                    continue;
                }
                
                var paragraphStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;

                if (paragraphStyle != null && paragraphStyle?.ToString()?.StartsWith("ART") == true)
                {
                    Console.WriteLine("[CTOR]\tZnaleziono artykuł w paragrafie: " + paragraph.InnerText);
                    Articles.Add(new Article(paragraph));
                }
            }
            Console.WriteLine("Znaleziono artykułów: " + Articles.Count);
        }

        public void RemoveSystemComments()
        {
            var commentPart = MainPart.WordprocessingCommentsPart;
            if (commentPart != null)
            {
                var comments = commentPart.Comments.Elements<Comment>().Where(c => c.Author == "System").ToList();
                foreach (var comment in comments)
                {
                    var commentId = comment.Id?.Value;
                    if (commentId == null) continue;

                    // Usuń zakresy komentarzy
                    var commentRangeStarts = MainPart.Document.Descendants<CommentRangeStart>().Where(c => c.Id == commentId).ToList();
                    var commentRangeEnds = MainPart.Document.Descendants<CommentRangeEnd>().Where(c => c.Id == commentId).ToList();
                    foreach (var rangeStart in commentRangeStarts)
                    {
                        rangeStart.Remove();
                    }
                    foreach (var rangeEnd in commentRangeEnds)
                    {
                        rangeEnd.Remove();
                    }

                    // Usuń odniesienia do komentarzy
                    var commentReferences = MainPart.Document.Descendants<CommentReference>().Where(c => c.Id == commentId).ToList();
                    foreach (var reference in commentReferences)
                    {
                        reference.Remove();
                    }

                    // Usuń komentarz
                    comment.Remove();
                }
            }
        }

        public int ParseHyperlinks()
        {
            int commentCount = 0;

            // Parsowanie paragrafów w treści dokumentu
            foreach (var paragraph in MainPart.Document.Descendants<Paragraph>())
            {
                commentCount += AddCommentsToHyperlinks(paragraph);
            }

            // Parsowanie przypisów dolnych
            if (MainPart.FootnotesPart != null)
            {
                foreach (var footnote in MainPart.FootnotesPart.Footnotes.Elements<Footnote>())
                {
                    foreach (var paragraph in footnote.Descendants<Paragraph>())
                    {
                        commentCount += AddCommentsToHyperlinks(paragraph);
                    }
                }
            }

            // Parsowanie przypisów końcowych
            if (MainPart.EndnotesPart != null)
            {
                foreach (var endnote in MainPart.EndnotesPart.Endnotes.Elements<Endnote>())
                {
                    foreach (var paragraph in endnote.Descendants<Paragraph>())
                    {
                        commentCount += AddCommentsToHyperlinks(paragraph);
                    }
                }
            }

            return commentCount;
        }

        private int AddCommentsToHyperlinks(Paragraph paragraph)
        {
            int commentCount = 0;
            foreach (var hyperlink in paragraph.Descendants<Hyperlink>())
            {
                string? hyperlinkUri = null;
                if (hyperlink.Id != null)
                {
                    var relationship = MainPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id);
                    if (relationship != null)
                    {
                        hyperlinkUri = relationship.Uri.ToString();
                    }
                }

                if (hyperlinkUri != null)
                {
                    var hyperlinkText = hyperlink.Descendants<Run>().Select(r => r.InnerText).FirstOrDefault();
                    var commentText = $"Hiperłącze: {hyperlinkUri}\nTekst: {hyperlinkText}";

                    Console.WriteLine("[HLINKS]\tDodawanie komentarza: " + commentText);

                    AddComment(hyperlink, commentText);

                    commentCount++;
                }
            }
            return commentCount;
        }

        private void AddComment(OpenXmlElement element, string commentText)
        {
            var commentPart = MainPart.WordprocessingCommentsPart;
            if (commentPart == null)
            {
                commentPart = MainPart.AddNewPart<WordprocessingCommentsPart>();
                commentPart.Comments = new Comments();
            }

            var commentId = commentPart.Comments.Elements<Comment>().Count().ToString();
            var comment = new Comment { Id = commentId, Author = "System", Date = DateTime.Now };
            comment.AppendChild(new Paragraph(new Run(new Text(commentText))));
            commentPart.Comments.Append(comment);

            var commentRangeStart = new CommentRangeStart { Id = commentId };
            var commentRangeEnd = new CommentRangeEnd { Id = commentId };

            if (element is Run)
            {
                element.InsertBefore(commentRangeStart, element.FirstChild);
                element.InsertAfter(commentRangeEnd, element.LastChild);
            }
            else if (element is Paragraph paragraph)
            {
                var firstRun = paragraph.Elements<Run>().FirstOrDefault();
                var lastRun = paragraph.Elements<Run>().LastOrDefault();

                if (firstRun != null)
                {
                    firstRun.InsertBeforeSelf(commentRangeStart);
                    firstRun.InsertAfterSelf(commentRangeEnd);
                }
                // else
                // {
                //     paragraph.InsertBefore(commentRangeStart, paragraph.FirstChild);
                // }

                // if (lastRun != null)
                // {
                //     lastRun.InsertAfterSelf(commentRangeEnd);
                // }
                // else
                // {
                //     paragraph.AppendChild(commentRangeEnd);
                // }
            }

            var commentReference = new CommentReference { Id = commentId };
            element.AppendChild(commentReference);
        }
 
        /// <summary>
        /// Cleans the properties of paragraphs within the main part of the document.
        /// This method performs the following actions:
        /// - Removes all BookmarkStart and BookmarkEnd elements.
        /// - Iterates through all paragraphs and processes each one:
        ///   - Clones the paragraph properties and retains only the paragraph style (pStyle).
        ///   - Removes all "rsid" attributes from runs.
        ///   - Removes all run properties except for styles (rStyle), vertical alignment (vertAlign), bold (b), and italic (i).
        ///   - Replaces the old paragraph with the new cleaned paragraph.
        /// </summary>
        internal void CleanParagraphProperties()
        {
            OpenXmlElement root = MainPart.Document;

            root.Descendants<BookmarkStart>().ToList().ForEach(b => b.Remove());
            root.Descendants<BookmarkEnd>().ToList().ForEach(b => b.Remove());

            var paragraphs = root.Descendants<Paragraph>().ToList();
            
            foreach (var paragraph in paragraphs)
            {
                Console.WriteLine("[CLEANING]\tPrzetwarzanie paragrafu: " + paragraph.InnerText);
                var newParagraph = new Paragraph();
                var newParagraphId = Guid.NewGuid().ToString("N").Substring(0, 16);
                newParagraph.ParagraphId = newParagraphId;

                var paragraphProperties = paragraph.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
                if (paragraphProperties != null)
                {
                    // Do nowego paragrafu przenieś tylko parametr pStyle
                    var pStyle = paragraphProperties.ParagraphStyleId;
                    paragraphProperties.RemoveAllChildren();
                    if (pStyle != null)
                    {
                        Console.WriteLine("[CLEANING]\tPrzenoszę styl paragrafu: " + pStyle.Val);
                        paragraphProperties.AppendChild(pStyle.CloneNode(true));
                    } else {
                        Console.WriteLine("[CLEANING]\tBrak stylu paragrafu!");
                        var firstRun = paragraph.Descendants<Run>().FirstOrDefault();
                        if (firstRun != null)
                            AddComment(firstRun, "Styl paragrafu nie zdefiniowany!");
                    }
                    newParagraph.ParagraphProperties = paragraphProperties;
                }

                foreach (var run in paragraph.Elements<Run>())
                {
                    // Usuń atrybuty rsid z runów
                    var rsidAttributes = run.GetAttributes().Where(a => a.LocalName.Contains("rsid")).ToList();
                    foreach (var rsidAttribute in rsidAttributes)
                    {
                        Console.WriteLine("[CLEANING]\tUsuwam atrybut: " + rsidAttribute.LocalName);
                        run.RemoveAttribute(rsidAttribute.LocalName, rsidAttribute.NamespaceUri);
                    }

                    // Usuń atrybuty poza stylami
                    var runProperties = run.RunProperties;
                    if (runProperties != null && runProperties.HasChildren)
                    {
                        var childrenToRemove = runProperties.Elements()
                            .Where(e => e.LocalName != "rStyle" && 
                                        e.LocalName != "vertAlign" && 
                                        e.LocalName != "b" && 
                                        e.LocalName != "i")
                            .ToList();
                        foreach (var child in childrenToRemove)
                        {
                            Console.WriteLine("[CLEANING]\tUsuwam element: " + child.OuterXml);
                            child.Remove();
                        }
                        if (!runProperties.HasChildren)
                        {
                            run.RunProperties = null;
                        }
                        // ReplaceFormattingWithStyle(run, runProperties);
                    }
                    newParagraph.AppendChild(run.CloneNode(true));
                }

                // Zamień stary paragraf na nowy
                paragraph.InsertAfterSelf(newParagraph);
                paragraph.Remove();
            }

            void ReplaceFormattingWithStyle(Run run, RunProperties runProperties)
            {
                if (runProperties.Elements<Italic>().Any())
                {
                    var rStyle = new RunStyle { Val = GetStyleID("_K_ - kursywa") };
                    runProperties.AppendChild(rStyle);
                    runProperties.Elements<Bold>().ToList().ForEach(e => e.Remove());
                    AddComment(run, "Zamieniono ręczne formatowanie kursywy na styl");
                } else if (runProperties.Elements<Bold>().Any())
                {
                    var rStyle = new RunStyle { Val = GetStyleID("_P_ - pogrubienie") };
                    runProperties.AppendChild(rStyle);
                    runProperties.Elements<Bold>().ToList().ForEach(e => e.Remove());
                    AddComment(run, "Zamieniono ręczne formatowanie pogrubienia na styl");
                }
                if (runProperties.Elements<VerticalTextAlignment>().Any() )
                {
                    var rStyle = new RunStyle { Val = GetStyleID("_IG_ - indeks górny") };
                    runProperties.AppendChild(rStyle);
                    runProperties.Elements<VerticalTextAlignment>().ToList().ForEach(e => e.Remove());
                    AddComment(run, "Zamieniono ręczne formatowanie indeksu górnego na styl");
                }
                if (runProperties.Elements<Bold>().Any())
                {
                    // AddComment(run, "Ręczne formatowanie pogrubienia");
                }
                if (runProperties.Elements<Italic>().Any())
                {
                    // AddComment(run, "Ręczne formatowanie kursywy");
                }
            }
        }

        public void MergeRuns()
        {
            var paragraphs = MainPart.Document.Descendants<Paragraph>()
                                                .Where(p => p.Elements<Run>().Count() > 1).ToList();

            foreach (var paragraph in paragraphs)
            {
                var runs = paragraph.Elements<Run>().ToList();
                Console.WriteLine("[RUN_MERGE]\tPrzetwarzanie paragrafu: " + paragraph.InnerText);
                Console.WriteLine("[RUN_MERGE]\tLiczba runów: " + runs.Count);
                Run newRun = null;

                foreach (var run in runs)
                {
                    if (run.RunProperties == null)
                    {
                        if (newRun == null)
                        {
                            newRun = new Run();
                        }
                        foreach (var child in run.Elements())
                        {
                            newRun.AppendChild(child.CloneNode(true));
                        }
                    }
                    else
                    {
                        if (newRun != null)
                        {
                            paragraph.AppendChild(newRun);
                            newRun = null;
                        }
                        paragraph.AppendChild(run.CloneNode(true));
                    }
                }

                if (newRun != null)
                {
                    paragraph.AppendChild(newRun);
                }

                // Remove all existing runs
                foreach (var run in runs)
                {
                    run.Remove();
                }
            }
        }

        internal void MergeTexts()
        {
            var runs = MainPart.Document.Descendants<Run>().Where(r => r.Elements<Text>().Count() > 1).ToList();

            foreach (var run in runs)
            {
                var newRun = new Run();
                Text? previousText = null;

                foreach (var element in run.Elements())
                {
                    if (element is Text textElement)
                    {
                        if (previousText == null)
                        {
                            previousText = new Text { Space = SpaceProcessingModeValues.Preserve, Text = textElement.Text };
                            newRun.AppendChild(previousText);
                        }
                        else
                        {
                            previousText.Text += textElement.Text;
                        }
                    }
                    else
                    {
                        newRun.AppendChild(element.CloneNode(true));
                        previousText = null;
                    }
                }

                run.InsertAfterSelf(newRun);
                run.Remove();
            }
        }

        // -------------

        private StringValue? GetStyleID(string styleName = "Normalny")
        {
            return MainPart.StyleDefinitionsPart?.Styles?.Descendants<Style>()
                                            .FirstOrDefault(s => s.StyleName?.Val == styleName)?.StyleId;
        }
        
        internal void Save()
        {
            _wordDoc.Save();
        }
        
        public void SaveAs(string newFilePath)
        {
            using (var newDoc = (WordprocessingDocument)_wordDoc.Clone(newFilePath))
            {
                newDoc.CompressionOption = CompressionOption.Maximum;
                newDoc.Save();
                //System.IO.Compression.ZipFile.ExtractToDirectory(newFilePath, Path.GetDirectoryName(newFilePath));
            }
        }
       
        internal void GenerateXML()
        {
            CustomXmlPart xmlPart = MainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml, "aktPrawny");
            var xmlDoc = new System.Xml.XmlDocument();
            var rootElement = xmlDoc.CreateElement(XMLConstants.Root);

            ProcessArticles(xmlDoc, rootElement, Articles);

            xmlDoc.AppendChild(rootElement);

            using (var stream = xmlPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                xmlDoc.Save(stream);
            }
            var xmlFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "LegalAct.xml");
            xmlDoc.Save(xmlFilePath);
            Console.WriteLine($"XML file saved at: {xmlFilePath}");
        }
    
        private void ProcessAmendments(XmlDocument xmlDoc, XmlElement parentElement, IEnumerable<Amendment> amendments, string amendmentType)
        {
            foreach (var amendment in amendments)
            {
                Console.WriteLine($"[XML]\t[{amendmentType}]\tPrzetwarzanie nowelizacji: " + amendment.Content);
                var amendmentElement = xmlDoc.CreateElement(XMLConstants.Amendment);
                amendmentElement.InnerText = amendment.Content;
                amendmentElement.SetAttribute("ustawaZmieniana", amendment.AmendedAct);
                parentElement.AppendChild(amendmentElement);
            }
        }

        private void ProcessArticles(XmlDocument xmlDoc, XmlElement rootElement, IEnumerable<Article> articles)
        {
            foreach (var article in articles)
            {
                Console.WriteLine("[XML]\t[ART]\tPrzetwarzanie artykułu: " + article.Content);
                var articleElement = xmlDoc.CreateElement(XMLConstants.Article);
                articleElement.SetAttribute(XMLConstants.Number, article.Number);
                articleElement.SetAttribute(XMLConstants.Amending, article.IsAmending ? "1" : "0");

                if (article.IsAmending)
                {
                    articleElement.SetAttribute(XMLConstants.PublicationYear, article.PublicationYear);
                    articleElement.SetAttribute(XMLConstants.PublicationNumber, article.PublicationNumber);
                }

                rootElement.AppendChild(articleElement);
                ProcessSubsections(xmlDoc, articleElement, article.Subsections);
            }
        }

        private void ProcessSubsections(XmlDocument xmlDoc, XmlElement parentElement, IEnumerable<Subsection> subsections)
        {
            foreach (var subsection in subsections)
            {
                Console.WriteLine("[XML]\t[UST]\tPrzetwarzanie ustepu: " + subsection.Content);
                var subsectionElement = xmlDoc.CreateElement(XMLConstants.Subsection);
                subsectionElement.InnerText = subsection.Content;
                subsectionElement.SetAttribute(XMLConstants.Number, subsection.Number.ToString());
                parentElement.AppendChild(subsectionElement);

                if (subsection.Amendments.Any())
                {
                    ProcessAmendments(xmlDoc, subsectionElement, subsection.Amendments, "ZUST");
                }

                ProcessPoints(xmlDoc, subsectionElement, subsection.Points);
            }
        }

        private void ProcessPoints(XmlDocument xmlDoc, XmlElement parentElement, IEnumerable<Point> points)
        {
            foreach (var point in points)
            {
                Console.WriteLine("[XML]\t[PKT]\tPrzetwarzanie punktu: " + point.Content);
                var pointElement = xmlDoc.CreateElement(XMLConstants.Point);
                pointElement.InnerText = point.Content;
                if (point.AmendmentOperations.Any())
                {
                    foreach (var amendmentOperation in point.AmendmentOperations)
                    {
                        var amendmentElement = xmlDoc.CreateElement("pn");
                        amendmentElement.SetAttribute("typ", amendmentOperation.Type.ToDescription());
                        //TODO - check for every target
                        var amendmentTarget = amendmentOperation.AmendmentTargets.FirstOrDefault();
                        if (amendmentTarget?.Article != null)
                        {
                            amendmentElement.SetAttribute("artykul", amendmentTarget.Article);
                        }
                        if (amendmentTarget?.Subsection != null)
                        {
                            amendmentElement.SetAttribute("ustep", amendmentTarget.Subsection);
                        }
                        if (amendmentTarget?.Point != null)
                        {
                            amendmentElement.SetAttribute("punkt", amendmentTarget.Point);
                        }
                        if (amendmentTarget?.Letter != null)
                        {
                            amendmentElement.SetAttribute("litera", amendmentTarget.Letter);
                        }
                        pointElement.AppendChild(amendmentElement);
                    }
                }

                pointElement.SetAttribute(XMLConstants.Number, point.Number);
                parentElement.AppendChild(pointElement);

                if (point.Amendments.Any())
                {
                    ProcessAmendments(xmlDoc, pointElement, point.Amendments, "ZPKT");
                }

                ProcessLetters(xmlDoc, pointElement, point.Letters);
            }
        }

        private void ProcessLetters(XmlDocument xmlDoc, XmlElement parentElement, IEnumerable<Letter> letters)
        {
            foreach (var letter in letters)
            {
                Console.WriteLine("[XML]\t[LIT]\tPrzetwarzanie litery: " + letter.Content);
                var letterElement = xmlDoc.CreateElement(XMLConstants.Letter);
                letterElement.InnerText = letter.Content;
                letterElement.SetAttribute(XMLConstants.LetterOrdinal, letter.Ordinal);
                parentElement.AppendChild(letterElement);

                if (letter.Amendments.Any())
                {
                    ProcessAmendments(xmlDoc, letterElement, letter.Amendments, "ZLIT");
                }

                ProcessTirets(xmlDoc, letterElement, letter.Tirets);
            }
        }
        
        private void ProcessTirets(XmlDocument xmlDoc, XmlElement parentElement, IEnumerable<Tiret> tirets)
        {
            foreach (var tiret in tirets)
            {
                Console.WriteLine("[XML]\t[TIR]\tPrzetwarzanie tiretu: " + tiret.Content);
                var tiretElement = xmlDoc.CreateElement(XMLConstants.Tiret);
                tiretElement.InnerText = tiret.Content;
                parentElement.AppendChild(tiretElement);
            }
        }

        internal void SaveAmendmentList()
        {
            var allAmendments = new List<string>();
            foreach (Article article in Articles)
            {
                if (article.AmendmentList != null && article.AmendmentList.Any())
                {
                    allAmendments.AddRange(article.AmendmentList);
                }
            }
            if (allAmendments.Any())
            {
                allAmendments = allAmendments.Distinct().ToList();
                var amendmentsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "nowele.txt");
                File.WriteAllLines(amendmentsFilePath, allAmendments);
                Console.WriteLine($"Amendments list saved at: {amendmentsFilePath}");
            }
            else
            {
                Console.WriteLine("No amendments found to save.");
            }
        }
    }    
}   