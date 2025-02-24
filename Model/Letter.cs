using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
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
}