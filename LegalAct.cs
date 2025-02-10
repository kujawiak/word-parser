using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO.Packaging;
using System.Linq;

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

        public void ValidateFormat()
        {
            int totalCommentCount = 0;
            Console.WriteLine("Sprawdzanie formatowania");

            foreach (var run in MainPart.Document.Descendants<Run>())
            {
                var runProperties = run.RunProperties;
                if (runProperties == null) continue;

                if (run.Descendants<FootnoteReference>().Any() || run.Descendants<EndnoteReference>().Any())
                {
                    Console.WriteLine("Pomijam przypis");
                    continue;
                }

                var bold = runProperties.Bold;
                if (bold != null && bold.Val != null && bold.Val == false)
                {
                    var commentText = $"Fragment tekstu nie jest pogrubiony: {run.InnerText}";
                    AddComment(run, commentText);
                    totalCommentCount++;
                }

                var styleId = runProperties.RunStyle?.Val?.Value;
                if (styleId == null)
                {
                    continue;
                }
                var styleName = MainPart.StyleDefinitionsPart.Styles.Descendants<StyleName>()
                    .FirstOrDefault(s => (s.Parent as Style)?.StyleId == styleId)?.Val?.Value;

                if (styleName == null)
                {
                    continue;
                }

                if (Program.VALID_STYLES_BOLD.Contains(styleName))
                {
                    if (runProperties.Bold == null || (runProperties.Bold.Val != null && runProperties.Bold.Val == false))
                    {
                        var commentText = $"Brak pogrubienia, wymaganego przez styl: {styleName}\nwe fragmencie: {run.InnerText}";
                        AddComment(run, commentText);
                        totalCommentCount++;
                    }
                }
            }

            Console.WriteLine($"Liczba dodanych komentarzy: {totalCommentCount}");
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

                    Console.WriteLine("Dodawanie komentarza: " + commentText);

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

        public void RemoveFontChanges(OpenXmlElement rootElement)
        {
            var elementsToRemove = rootElement.Descendants<RunFonts>()
                                            .Where(rf => rf.ComplexScript != null
                                            || rf.HighAnsi != null
                                            || rf.Ascii != null)
                                            .ToList();

            // Usuń znalezione elementy
            foreach (var element in elementsToRemove)
            {
                System.Console.WriteLine("Wykryta zmiana czcionki: " + element.OuterXml);
                // var run = element.Ancestors<Run>().FirstOrDefault();
                // if (run != null)
                // {
                //     AddComment(run, "Usunięto zmianę czcionki: " + element.OuterXml);
                // }
                element.Remove();   
            }

            // Rekurencyjnie usuń elementy w pod-elementach
            foreach (var child in rootElement.Elements())
            {
                RemoveFontChanges(child);
            }
        }

        internal void RemoveRsidAttributes(OpenXmlElement rootElement)
        {
            var elementsWithRsid = rootElement.Descendants()
                .Where(e => e.HasAttributes)
                .ToList();

            foreach (var element in elementsWithRsid)
            {
                element.RemoveAttribute("rsidRPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                element.RemoveAttribute("rsidR", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                element.RemoveAttribute("rsidP", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                element.RemoveAttribute("rsidDel", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                element.RemoveAttribute("rsidTr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                element.RemoveAttribute("rsidRDefault", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                element.RemoveAttribute("rsidSect", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            }

            // Rekurencyjnie usuń atrybuty w pod-elementach
            foreach (var child in rootElement.Elements())
            {
                RemoveRsidAttributes(child);
            }
        }
        
        internal void CleanRunProperties(OpenXmlElement rootElement)
        {
            // Znajdź wszystkie elementy <w:rPr> wewnątrz <w:r>
            var runProperties = rootElement.Descendants<Run>()
                .Select(r => r.RunProperties)
                .Where(rPr => rPr != null && rPr.HasChildren)
                .ToList();

            // Wyczyść dzieci <w:rPr>
            foreach (var rPr in runProperties)
            {
                var childrenToRemove = rPr.Elements()
                                    .Where(child => child.LocalName != "rStyle").ToList(); 
                foreach (var child in childrenToRemove) 
                { 
                    switch (child.LocalName) 
                    { 
                        case "b": 
                            AddComment(rPr.Ancestors<Run>().FirstOrDefault(), "Usunięto formatowanie pogrubienia");
                            child.Remove();
                            break; 
                        case "i": 
                            AddComment(rPr.Ancestors<Run>().FirstOrDefault(), "Usunięto formatowanie kursywy");
                            child.Remove(); 
                            break; 
                        case "smallCaps": 
                            AddComment(rPr.Ancestors<Run>().FirstOrDefault(), "Usunięto formatowanie małych liter");
                            child.Remove(); 
                            break; 
                        case "caps": 
                            AddComment(rPr.Ancestors<Run>().FirstOrDefault(), "Usunięto formatowanie wielkich liter");
                            child.Remove(); 
                            break; 
                        case "color": 
                            AddComment(rPr.Ancestors<Run>().FirstOrDefault(), "Usunięto formatowanie koloru tekstu");
                            child.Remove(); 
                            break; 
                        case "highlight": 
                            AddComment(rPr.Ancestors<Run>().FirstOrDefault(), "Usunięto formatowanie koloru tła");
                            child.Remove(); 
                            break;
                        case "rFonts":
                            AddComment(rPr.Ancestors<Run>().FirstOrDefault(), "Usunięto formatowanie czcionki");
                            child.Remove(); 
                            break;
                        default:
                            System.Console.WriteLine("Usuwam element: " + child.OuterXml);
                            AddComment(rPr.Ancestors<Run>().FirstOrDefault(), "Usunięto element: " + child.OuterXml);
                            child.Remove();
                            break;
                    }
                    // AddComment(rPr, "Usuwam element: " + child.OuterXml);
                }
            }

            // Rekurencyjnie przetwórz pod-elementy
            foreach (var child in rootElement.Elements())
            {
                CleanRunProperties(child);
            }
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
// -----------

public void MergeAdjacentRunsWithSameFormatting()
{
    var paragraphs = MainPart.Document.Descendants<Paragraph>().ToList();

    foreach (var paragraph in paragraphs)
    {
        System.Console.WriteLine("Przetwarzanie paragrafu: " + paragraph.InnerText);
        var newParagraph = new Paragraph(paragraph.ParagraphProperties?.CloneNode(true));
        Run currentRun = null;

        foreach (var element in paragraph.Elements())
        {
            System.Console.WriteLine($"Przetwarzanie elementu <{element.LocalName}>: {element.InnerText}");

            if (element is Run run)
            {
                if (currentRun == null || !AreRunPropertiesEqual(currentRun.RunProperties, run.RunProperties))
                {
                    // Utwórz nowy Run i skopiuj właściwości tylko raz
                    currentRun = new Run(run.RunProperties?.CloneNode(true));
                    newParagraph.AppendChild(currentRun);
                }
                
                // Dodaj elementy tekstu do bieżącego Run, unikając ponownego tworzenia <w:rPr>
                foreach (var child in run.Elements().Where(c => !(c is RunProperties)))
                {
                    if (currentRun.LastChild is Text lastText && child is Text textChild)
                    {
                        
                        if (lastText.Space == null)
                        {
                            System.Console.WriteLine("Łączę teksty: " + lastText.Text + " + " + textChild.Text);
                            lastText.Text += textChild.Text;
                        }
                        else
                        {
                            currentRun.AppendChild(textChild.CloneNode(true));
                        }
                    }
                    else
                    {
                        currentRun.AppendChild(child.CloneNode(true));
                    }
                }
            }
            else if (element is ParagraphProperties)
            {
                newParagraph.RemoveAllChildren<ParagraphProperties>();
                newParagraph.Append(element.CloneNode(true));
            }
            else
            {
                System.Console.WriteLine("Nieznany element: " + element.LocalName);
            }
        }

        paragraph.InsertAfterSelf(newParagraph);
        System.Console.WriteLine("Nowy paragraf: " + newParagraph.InnerText);
        paragraph.Remove();
    }
}

static bool AreRunPropertiesEqual(RunProperties rp1, RunProperties rp2)
{
    if (rp1 == null && rp2 == null)
        return true;

    if (rp1 == null || rp2 == null)
        return false;

    return rp1.OuterXml == rp2.OuterXml;
}
// -------------
        internal void ClearTrackingChanges()
        {
            //throw new NotImplementedException();
            if (this.SettingsPart != null)
            {
                var settings = this.SettingsPart.Settings;
                var rsidList = settings.Descendants<Rsid>()
                            .Where(rsid => rsid.Val != null)
                            .Select(rsid => rsid.Val!.Value).ToList();

                string[] trackingAttributes = { "rsidR", "rsidRPr", "rsidDel", "rsidP", "rsidRDefault" };

                foreach (var element in MainPart.Document.Descendants())
                {
                    foreach (var attribute in trackingAttributes)
                    {
                        var rsidAttribute = element.GetAttributes().FirstOrDefault(a => a.LocalName == attribute);
                        if (rsidAttribute != null && rsidList.Contains(rsidAttribute.Value))
                        {
                            element.RemoveAttribute(rsidAttribute.LocalName, rsidAttribute.NamespaceUri);
                        }
                    }
                }

                var rsidsElement = settings.Descendants<Rsid>().FirstOrDefault();
                rsidsElement?.Remove();
            }
            else
            {
                System.Console.WriteLine("Brak ustawień dokumentu.");
            }
        }

        internal void RemovePreserveAttributes(OpenXmlElement rootElement)
        {
            // Usuń elementy <w:bookmarkStart> i <w:bookmarkEnd>
            rootElement.Descendants<BookmarkStart>().ToList().ForEach(b => b.Remove());
            rootElement.Descendants<BookmarkEnd>().ToList().ForEach(b => b.Remove());

            var paragraphs = rootElement.Descendants<Paragraph>().ToList();
            
            foreach (var paragraph in paragraphs)
            {
                System.Console.WriteLine("Przetwarzanie paragrafu: " + paragraph.InnerText);
                var newParagraph = new Paragraph();

                var paragraphProperties = paragraph.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
                if (paragraphProperties != null)
                {
                    var pStyle = paragraphProperties.ParagraphStyleId;
                    paragraphProperties.RemoveAllChildren();
                    if (pStyle != null)
                    {
                    System.Console.WriteLine("Przenoszę styl paragrafu: " + pStyle.Val);
                    paragraphProperties.AppendChild(pStyle.CloneNode(true));
                    }
                    else
                    {
                    System.Console.WriteLine("Brak stylu paragrafu!");
                    var firstRun = paragraph.Descendants<Run>().FirstOrDefault();
                    if (firstRun != null)
                        AddComment(firstRun, "Styl paragrafu nie zdefiniowany!");
                    }
                    newParagraph.ParagraphProperties = paragraphProperties;
                }

                foreach (var run in paragraph.Elements<Run>())
                {
                    var rsidAttributes = run.GetAttributes().Where(a => a.LocalName.Contains("rsid")).ToList();
                    foreach (var rsidAttribute in rsidAttributes)
                    {
                        System.Console.WriteLine("Usuwam atrybut: " + rsidAttribute.LocalName);
                        run.RemoveAttribute(rsidAttribute.LocalName, rsidAttribute.NamespaceUri);
                    }

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
                            System.Console.WriteLine("Usuwam element: " + child.OuterXml);
                            child.Remove();
                        }
                        // ReplaceFormattingWithStyle(run, runProperties);
                    }
                    newParagraph.AppendChild(run.CloneNode(true));
                }

                paragraph.InsertAfterSelf(newParagraph);
                paragraph.Remove();
            }
        }

            void ReplaceFormattingWithStyle(Run run, RunProperties runProperties)
            {
                bool isBold = runProperties.Elements<Bold>().Any();
                bool isItalic = runProperties.Elements<Italic>().Any();
                bool isSuperscript = runProperties.Elements<VerticalTextAlignment>().Any(v => v.Val == VerticalPositionValues.Superscript);
                bool isSubscript = runProperties.Elements<VerticalTextAlignment>().Any(v => v.Val == VerticalPositionValues.Subscript);

                if (isSuperscript)
                {
                    if (isBold && isItalic)
                    {
                    ApplyStyle(run, "_IG_P_K_ - indeks górny i pogrubienie kursywa");
                    }
                    else if (isBold)
                    {
                    ApplyStyle(run, "_IG_P_ - indeks górny i pogrubienie");
                    }
                    else if (isItalic)
                    {
                    ApplyStyle(run, "_IG_K_ - indeks górny i kursywa");
                    }
                    else
                    {
                    ApplyStyle(run, "_IG_ - indeks górny");
                    }
                }
                else if (isSubscript)
                {
                    if (isBold && isItalic)
                    {
                    ApplyStyle(run, "_ID_P_K_ - indeks dolny i pogrubienie kursywa");
                    }
                    else if (isBold)
                    {
                    ApplyStyle(run, "_ID_P_ - indeks dolny i pogrubienie");
                    }
                    else if (isItalic)
                    {
                    ApplyStyle(run, "_ID_K_ - indeks dolny i kursywa");
                    }
                    else
                    {
                    ApplyStyle(run, "_ID_ - indeks dolny");
                    }
                }
                else
                {
                    if (isBold && isItalic)
                    {
                    ApplyStyle(run, "_P_K_ - pogrubienie kursywa");
                    }
                    else if (isBold)
                    {
                    ApplyStyle(run, "_P_ - pogrubienie");
                    }
                    else if (isItalic)
                    {
                    ApplyStyle(run, "_K_ - kursywa");
                    }
                }

                runProperties.Elements<Bold>().ToList().ForEach(e => e.Remove());
                runProperties.Elements<Italic>().ToList().ForEach(e => e.Remove());
                runProperties.Elements<VerticalTextAlignment>().ToList().ForEach(e => e.Remove());
            }

            void ApplyStyle(Run run, string styleName)
            {
                var rStyle = new RunStyle { Val = GetStyleID(styleName) };
                run.RunProperties.AppendChild(rStyle);
                AddComment(run, $"Zamieniono ręczne formatowanie na styl: {styleName}");
            }
        
        private StringValue? GetStyleID(string styleName = "Normalny")
        {
            return MainPart.StyleDefinitionsPart?.Styles?.Descendants<Style>()
                                            .FirstOrDefault(s => s.StyleName?.Val == styleName)?.StyleId;
        }
    }
}