using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser
{
    public class BaseEntity {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Content { get; set; }
        public Paragraph? Paragraph { get; set; }
        public BaseEntity(Paragraph paragraph)
        {
            Paragraph = paragraph;
            Content = paragraph.InnerText.Sanitize();
        }
    }

    // Tytuł
    public class Title : BaseEntity {
        public List<Part> Parts { get; set; } = new List<Part>();

        public Title(Paragraph paragraph) : base(paragraph)
        {
        }
    }

    // Dział 
    public class Part {
        public List<Chapter> Chapters { get; set; }
    }

    // Rozdział
    public class Chapter {
        public List<Section> Sections { get; set; }
    }

    // Oddział
    public class Section {
        public List<Article> Articles { get; set; }

        internal void AddArticle(Paragraph paragraph)
        {
            var article = new Article(paragraph);
            Articles.Add(article);
        }
    }

    // Artykuł
    public class Article : BaseEntity {
        public string Number { get; set; }
        public bool IsAmending { get; set; }
        public string? PublicationYear { get; set; }
        public string? PublicationNumber { get; set; }
        public List<Subsection> Subsections { get; set; }

        public Article(Paragraph paragraph) : base(paragraph)
        {
            Number = SetNumber();
            IsAmending = SetAmendment();
            Subsections = [new Subsection(paragraph, this)];
            var ordinal = 1;
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("ART") != true)
            {
                ordinal++;
                if (nextParagraph.StyleId("UST") == true)
                {
                    Subsections.Add(new Subsection(nextParagraph, this, ordinal));
                }
                paragraph = nextParagraph;
            }
        }

        string SetNumber()
        {
            var match = Regex.Match(Content, @"Art\. ([\w\d]+)\.");
            return match.Success ? match.Groups[1].Value : "Unknown";
        }

        bool SetAmendment()
        {
            var publication = new Regex(@"Dz\.\sU\.\sz\s(\d{4})\sr\.\spoz\.\s(\d+)");
            if (publication.Match(Content).Success)
            {
                PublicationYear = publication.Match(Content).Groups[1].Value;
                PublicationNumber = publication.Match(Content).Groups[2].Value;
                return true;
            } else {
                return false;
            }
        }
    }

    // Ustęp
    public class Subsection : BaseEntity {
        public Article Parent { get; set; }
        public List<Point> Points { get; set; }
        public int Number { get; set; }

        public Subsection(Paragraph paragraph, Article parent, int ordinal = 1) : base(paragraph)
        {
            Parent = parent;
            Number = ordinal;
            Points = new List<Point>();
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("UST") != true
                    && nextParagraph.StyleId("ART") != true)
            {
                if (nextParagraph.StyleId("PKT") == true)
                {
                    Points.Add(new Point(nextParagraph, this));
                }
                paragraph = nextParagraph;
            }
        }
    }

    // Punkt
    public class Point : BaseEntity {
        public Subsection Parent { get; set; }
        public List<Letter> Letters { get; set; }
        public string Number { get; set; }
        public Point(Paragraph paragraph, Subsection parent) : base(paragraph)
        {
            Parent = parent;
            Number = Content.ExtractOrdinal();
            Letters = new List<Letter>();
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("PKT") != true
                    && nextParagraph.StyleId("UST") != true
                    && nextParagraph.StyleId("ART") != true)
            {
                if (nextParagraph.StyleId("LIT") == true)
                {
                    Letters.Add(new Letter(nextParagraph, this));
                }
                paragraph = nextParagraph;
            }
        }
    }

    // Litera
    public class Letter : BaseEntity {
        public Point Parent { get; set; }
        public List<Tiret> Tirets { get; set; }
        public string Ordinal { get; set; }

        public Letter(Paragraph paragraph, Point parent) : base(paragraph)
        {
            Parent = parent;
            Ordinal = Content.ExtractOrdinal();
            Tirets = new List<Tiret>();
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("LIT") != true
                    && nextParagraph.StyleId("PKT") != true
                    && nextParagraph.StyleId("UST") != true
                    && nextParagraph.StyleId("ART") != true)
            {
                if (nextParagraph.StyleId("TIRET") == true)
                {
                    Tirets.Add(new Tiret(nextParagraph));
                }
                paragraph = nextParagraph;
            }
        }
    }

    // Tiret
    public class Tiret : BaseEntity {
        public Tiret(Paragraph paragraph) : base(paragraph)
        {
        }
    }
}