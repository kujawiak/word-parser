using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
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
}