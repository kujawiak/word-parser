using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
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
}