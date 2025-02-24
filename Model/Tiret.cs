using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
        public class Tiret : BaseEntity, IAmendable {
        Letter Parent { get; set; }
        public int Number { get; set; }
        public List<Amendment> Amendments { get; set; }
        public Tiret(Paragraph paragraph, Letter parent, int ordinal = 1) : base(paragraph)
        {
            Article = parent.Article;
            Subsection = parent.Subsection;
            Point = parent.Point;
            Letter = parent;
            Parent = parent;
            Number = ordinal;
            Amendments = new List<Amendment>();
            bool isAdjacent = true;
            while (paragraph.NextSibling() is Paragraph nextParagraph 
                    && nextParagraph.StyleId("TIR") != true
                    && nextParagraph.StyleId("LIT") != true
                    && nextParagraph.StyleId("PKT") != true
                    && nextParagraph.StyleId("UST") != true
                    && nextParagraph.StyleId("ART") != true)
            {
                if (nextParagraph.StyleId("Z") == true && isAdjacent == true)
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