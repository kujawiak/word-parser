using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
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
            if (Article.IsAmending && Content.Contains("uchyla się"))
            {
                Amendments.Add(new Amendment(paragraph, this));
            }
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
}