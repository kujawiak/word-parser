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
            Tiret = parent.Tiret ?? (parent as Tiret);
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
                var tir = Tiret?.Content;
                var parts = new List<string>
                {
                    Article?.PublicationNumber?.ToString() ?? string.Empty,
                    Article?.PublicationYear?.ToString() ?? string.Empty
                };
                // if (!string.IsNullOrEmpty(ust)) parts.Add(ust);
                if (!string.IsNullOrEmpty(pkt)) parts.Add(pkt);
                if (!string.IsNullOrEmpty(lit)) parts.Add(lit);
                if (!string.IsNullOrEmpty(tir)) parts.Add(tir);
                var regexInput = parts.Count > 0 ? string.Join("|", parts) : null;
                Parent.Article.AmendmentList.Add(regexInput);
                return regexInput.GetAmendingProcedure();
            }
        }
    }
}