using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser
{
    public class BaseEntity 
    {
        public Article? Article { get; set; }
        public Subsection? Subsection { get; set; }
        public Point? Point { get; set; }
        public Letter? Letter { get; set; }
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Content { get; set; }
        public Paragraph? Paragraph { get; set; }
        public BaseEntity(Paragraph paragraph)
        {
            Paragraph = paragraph;
            Content = paragraph.InnerText.Sanitize();
        }
    }

    public interface IAmendable
    {
        List<Amendment> Amendments { get; set; }
    }

    // Tytuł
    public class Title : BaseEntity 
    {   
        public string TitleText { get; set; }
        public List<Part> Parts { get; set; } = new List<Part>();

        public Title(Paragraph paragraph) : base(paragraph)
        {
        }
    }

    // Dział 
    public class Part 
    {
        Title Parent { get; set; }
        public string Number { get; set; }
        public List<Chapter> Chapters { get; set; }
    }

    // Rozdział
    public class Chapter 
    {
        Part Parent { get; set; }
        public string Number { get; set; }
        public List<Section> Sections { get; set; }
    }

    // Oddział
    public class Section 
    {
        Chapter Parent { get; set; }
        public string Number { get; set; }
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
        public List<string> AmendmentList { get; set; }

        public Article(Paragraph paragraph) : base(paragraph)
        {
            Number = SetNumber();
            IsAmending = SetAmendment();
            Subsections = [new Subsection(paragraph, this)];
            AmendmentList = new List<string>();
            var ordinal = 1;
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("ART") != true)
            {
                if (nextParagraph.StyleId("UST") == true)
                {
                    ordinal++;
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
    public class Subsection : BaseEntity, IAmendable {
        public Article Parent { get; set; }
        public List<Point> Points { get; set; }
        public int Number { get; set; }
        public List<Amendment> Amendments { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Subsection"/> class.
        /// </summary>
        /// <param name="paragraph">The paragraph associated with this subsection.</param>
        /// <param name="article">The parent article of this subsection.</param>
        /// <param name="ordinal">The ordinal number of this subsection. Default is 1.</param>
        public Subsection(Paragraph paragraph, Article article, int ordinal = 1) : base(paragraph)
        {
            Parent = article;
            Article = article;
            Number = ordinal;
            Points = new List<Point>();
            Amendments = new List<Amendment>();
            bool isAdjacent = true;
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("UST") != true
                    && nextParagraph.StyleId("ART") != true)
            {
                if (nextParagraph.StyleId("PKT") == true)
                {
                    Points.Add(new Point(nextParagraph, this));
                    isAdjacent = false;
                }
                else if (nextParagraph.StyleId("Z") == true && isAdjacent)
                {
                    Amendments.Add(new Amendment(nextParagraph, this));
                }
                else 
                {
                    isAdjacent = false;
                }
                paragraph = nextParagraph;
            }
        }
    }

    // Punkt
    public class Point : BaseEntity, IAmendable {
        public Subsection Parent { get; set; }
        public List<Letter> Letters { get; set; }
        public List<Amendment> Amendments { get; set; }
        public string Number { get; set; }
        public Point(Paragraph paragraph, Subsection parent) : base(paragraph)
        {
            Article = parent.Parent;
            Subsection = parent;
            Parent = parent;
            Subsection = parent;
            Number = Content.ExtractOrdinal();
            Letters = new List<Letter>();
            Amendments = new List<Amendment>();
            bool isAdjacent = true;
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("PKT") != true
                    && nextParagraph.StyleId("UST") != true
                    && nextParagraph.StyleId("ART") != true)
            {
                if (nextParagraph.StyleId("LIT") == true)
                {
                    Letters.Add(new Letter(nextParagraph, this));
                    isAdjacent = false;
                }
                else if (nextParagraph.StyleId("Z") == true && isAdjacent == true)
                {
                    Amendments.Add(new Amendment(nextParagraph, this));
                }
                else 
                {
                    isAdjacent = false;
                }
                paragraph = nextParagraph;
            }
        }
    }

    // Litera
    public class Letter : BaseEntity, IAmendable {
        public Point Parent { get; set; }
        public List<Tiret> Tirets { get; set; }
        public string Ordinal { get; set; }
        public List<Amendment> Amendments { get; set; }

        public Letter(Paragraph paragraph, Point parent) : base(paragraph)
        {
            Article = parent.Article;
            Subsection = parent.Subsection;
            Point = parent;
            Parent = parent;
            Ordinal = Content.ExtractOrdinal();
            Tirets = new List<Tiret>();
            Amendments = new List<Amendment>();
            bool isAdjacent = true;
            var tiretCount = 1;
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("LIT") != true
                    && nextParagraph.StyleId("PKT") != true
                    && nextParagraph.StyleId("UST") != true
                    && nextParagraph.StyleId("ART") != true)
            {
                if (nextParagraph.StyleId("TIRET") == true)
                {
                    Tirets.Add(new Tiret(nextParagraph, this, tiretCount));
                    tiretCount++;
                }
                else if (nextParagraph.StyleId("Z") == true && isAdjacent == true)
                {
                    Amendments.Add(new Amendment(nextParagraph, this));
                }
                else 
                {
                    isAdjacent = false;
                }
                paragraph = nextParagraph;
            }
        }
    }

    // Tiret
    public class Tiret : BaseEntity {
        Letter Parent { get; set; }
        public int Number { get; set; }
        public Tiret(Paragraph paragraph, Letter parent, int ordinal = 1) : base(paragraph)
        {
            Article = parent.Article;
            Subsection = parent.Subsection;
            Point = parent.Point;
            Letter = parent;
            Parent = parent;
            Number = ordinal;
        }
    }

    public class Amendment : BaseEntity
    {
        public BaseEntity Parent { get; set; }
        public Amendment(Paragraph paragraph, BaseEntity parent) : base(paragraph)
        {
            Article = parent.Article ?? (parent as Article);
            Subsection = parent.Subsection ?? (parent as Subsection);
            Point = parent.Point ?? (parent as Point);
            Letter = parent.Letter ?? (parent as Letter);
            Parent = parent;
            Paragraph = paragraph;
        }

        public string? AmendedAct { 
            get
            {
                var art = Article?.Content;
                var ust = Subsection?.Content;
                var pkt = Point?.Content;
                var lit = Letter?.Content;
                var parts = new List<string>();
                if (!string.IsNullOrEmpty(art)) parts.Add(art);
                //if (!string.IsNullOrEmpty(ust)) parts.Add(ust);
                if (!string.IsNullOrEmpty(pkt)) parts.Add(pkt);
                if (!string.IsNullOrEmpty(lit)) parts.Add(lit);
                var regexInput = parts.Count > 0 ? string.Join("|", parts) : null;
                Parent.Article.AmendmentList.Add(regexInput);
                return regexInput.GetAmendingProcedure();
            }
        }
    }
}