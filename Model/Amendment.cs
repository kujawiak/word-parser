using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
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