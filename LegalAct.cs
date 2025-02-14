using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO.Packaging;
using System.Linq;
using System.Xml;

namespace WordParser
{
    public class LegalAct
    {
        public WordprocessingDocument _wordDoc
         { get; }
        public MainDocumentPart MainPart { get; }
        public DocumentSettingsPart? SettingsPart { get; }

        public LegalAct(WordprocessingDocument wordDoc)
        {
            _wordDoc = wordDoc;
            MainPart = _wordDoc.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is null.");
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
            }
        }
       
        internal void GenerateXMLSchema()
        {
            var xmlPart = MainPart.AddNewPart<CustomXmlPart>("application/xml", "rIdLegalActStructure");
            var xmlDoc = new System.Xml.XmlDocument();
            var rootElement = xmlDoc.CreateElement("ustawa");

            foreach (var paragraph in MainPart.Document.Descendants<Paragraph>()
                                                        .Where(p => p.InnerText.StartsWith("Art."))
                                                        .ToList())
            {
                if (paragraph.ParagraphProperties == null)
                {
                    Console.WriteLine("[XML]\t[ART]\tBrak właściwości paragrafu!");
                    continue;
                }
                if (paragraph.ParagraphProperties.ParagraphStyleId == null)
                {
                    Console.WriteLine("[XML]\t[ART]\tBrak stylu paragrafu!");
                    continue;
                }
                
                var paragraphStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;

                if (paragraphStyle.ToString().StartsWith("ART"))
                {
                    Console.WriteLine("[XML]\t[ART]\tPrzetwarzanie paragrafu/artykułu: " + paragraph.InnerText);
                    var isAmending = false;
                    var paragraphText = System.Text.RegularExpressions.Regex.Replace(paragraph.InnerText, @"\s+", " ");
                    var match = System.Text.RegularExpressions.Regex.Match(paragraphText, @"Art\. ([\w\d]+)\.");
                    var articleNumber = match.Success ? match.Groups[1].Value : "Unknown";

                    var articleElement = xmlDoc.CreateElement("artykul");
                    articleElement.SetAttribute("numer", articleNumber);
                    articleElement.SetAttribute("paraId", paragraph.ParagraphId?.ToString() ?? "Unknown");

                    var dzURegex = new System.Text.RegularExpressions.Regex(@"Dz\.\sU\.\sz\s(\d{4})\sr\.\spoz\.\s(\d+)");
                    var dzUMatch = dzURegex.Match(paragraphText);
                    if (dzUMatch.Success)
                    {
                        isAmending = true;
                        articleElement.SetAttribute("rok_publikatora", dzUMatch.Groups[1].Value);
                        articleElement.SetAttribute("numer_publikatora", dzUMatch.Groups[2].Value);
                    }

                    articleElement.InnerText = paragraph.InnerText;
                    
                    rootElement.AppendChild(articleElement);
                    GenerateXMLSchemaForAmendingPart(paragraph, articleElement, isAmending);
                    GenerateXMLSchemaForArticle(paragraph, articleElement, isAmending);
                }
            }

            xmlDoc.AppendChild(rootElement);

            using (var stream = xmlPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                xmlDoc.Save(stream);
            }
        }

        private void GenerateXMLSchemaForAmendingPart(Paragraph nextParagraph, XmlElement rootElement, bool isAmending)
        {
            var nextParagraphs = nextParagraph.ElementsAfter().OfType<Paragraph>().ToList();
            foreach (var paragraph in nextParagraphs)
            {
                System.Console.WriteLine("[XML]\t[Z]\tPrzetwarzanie paragrafu: " + paragraph.InnerText);
                var paragraphStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
                if (paragraphStyle != null)
                {
                    if (!paragraphStyle.ToString().StartsWith("Z"))
                    {
                        break;
                    }
                    var amendingElement = rootElement.OwnerDocument.CreateElement("zmiana_nowelizacyjna");
                    amendingElement.InnerText = paragraph.InnerText;
                    rootElement.AppendChild(amendingElement);
                }
            }
        }

        private void GenerateXMLSchemaForArticle(Paragraph paragraph, XmlElement articleElement, bool isAmending)
        {
            var nextParagraph = paragraph.NextSibling<Paragraph>();
            while (nextParagraph != null && !nextParagraph.InnerText.StartsWith("Art."))
            {
                System.Console.WriteLine("[XML]\t[UST/PKT]\tPrzetwarzanie paragrafu: " + nextParagraph.InnerText);
                var paragraphStyle = nextParagraph.ParagraphProperties?.ParagraphStyleId?.Val;
                if (paragraphStyle != null)
                {
                    if (paragraphStyle.ToString().StartsWith("UST"))
                    {
                        var sectionElement = articleElement.OwnerDocument.CreateElement("ustep");
                        sectionElement.InnerText = nextParagraph.InnerText;
                        articleElement.AppendChild(sectionElement);
                    } else if (paragraphStyle.ToString().StartsWith("PKT"))
                    {
                        var pointElement = articleElement.OwnerDocument.CreateElement("punkt");
                        pointElement.InnerText = nextParagraph.InnerText;
                        pointElement.SetAttribute("numer", nextParagraph.InnerText.Split(')')[0]);
                        pointElement.SetAttribute("paraId", nextParagraph.ParagraphId?.ToString() ?? "Unknown");
                        
                        articleElement.AppendChild(pointElement);
                        GenerateXMLSchemaForAmendingPart(nextParagraph, pointElement, isAmending);
                        GenerateXMLSchemaForPoint(nextParagraph, pointElement, isAmending);
                    }
                }
                nextParagraph = nextParagraph.NextSibling<Paragraph>();
            }
        }

        private void GenerateXMLSchemaForPoint(Paragraph paragraph, XmlElement pointElement, bool isAmending)
        {
            var nextParagraph = paragraph.NextSibling<Paragraph>();
            while (nextParagraph != null && !(nextParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.ToString().StartsWith("PKT") == true))
            {
                System.Console.WriteLine("[XML]\t[LIT]\tPrzetwarzanie paragrafu: " + nextParagraph.InnerText);
                var paragraphStyle = nextParagraph.ParagraphProperties?.ParagraphStyleId?.Val;
                if (paragraphStyle != null)
                {
                    if (paragraphStyle.ToString().StartsWith("LIT"))
                    {
                        var letterElement = pointElement.OwnerDocument.CreateElement("litera");
                        letterElement.InnerText = nextParagraph.InnerText;
                        letterElement.SetAttribute("lit", nextParagraph.InnerText.Split(')')[0]);
                        letterElement.SetAttribute("paraId", nextParagraph.ParagraphId?.ToString() ?? "Unknown");
                        
                        pointElement.AppendChild(letterElement);
                        GenerateXMLSchemaForAmendingPart(nextParagraph, letterElement, isAmending);
                        GenerateXMLSchemaForLetter(nextParagraph, letterElement, isAmending);
                    }
                }
                nextParagraph = nextParagraph.NextSibling<Paragraph>();
            }
        }

        private void GenerateXMLSchemaForLetter(Paragraph nextParagraph, XmlElement letterElement, bool isAmending)
        {
            var nextParagraphs = nextParagraph.ElementsAfter().OfType<Paragraph>().ToList();
            foreach (var paragraph in nextParagraphs)
            {
                System.Console.WriteLine("[XML]\t[TIR]\tPrzetwarzanie paragrafu: " + paragraph.InnerText);
                var paragraphStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
                if (paragraphStyle != null)
                {
                    if (paragraphStyle.ToString().StartsWith("ART") || 
                        paragraphStyle.ToString().StartsWith("UST") || 
                        paragraphStyle.ToString().StartsWith("PKT"))
                    {
                        break;
                    }
                    if (paragraphStyle.ToString().StartsWith("TIR"))
                    {
                        var sectionElement = letterElement.OwnerDocument.CreateElement("tiret");
                        sectionElement.InnerText = paragraph.InnerText;
                        letterElement.AppendChild(sectionElement);

                        GenerateXMLSchemaForAmendingPart(paragraph, letterElement, isAmending);
                    }
                }
            }
        }
    }
}