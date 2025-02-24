using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace WordParser.Model
{
    public class BaseEntity 
    {
        public Article? Article { get; set; }
        public Subsection? Subsection { get; set; }
        public Point? Point { get; set; }
        public Letter? Letter { get; set; }
        public Tiret? Tiret { get; set; }
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Content { get; set; }
        public Paragraph? Paragraph { get; set; }
        public List<AmendmentOperation> AmendmentOperations { get; set; } = new List<AmendmentOperation>();
        public BaseEntity(Paragraph paragraph)
        {
            Paragraph = paragraph;
            Content = paragraph.InnerText.Sanitize();
        }

        internal void TryParseAmendingOperation()
        {
            if (Content.Contains("uchyla się"))
            {
                var amendmentOperation = new AmendmentOperation
                {
                    Type = AmendmentOperationType.Repealed
                };
                var regex = new .Regex(@"art\. (\d+[\w\-]*)");
                var matches = regex.Matches(Content);
                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    var articleNumber = match.Groups[1].Value;
                    amendmentOperation.AmendmentTargets.Add(new AmendmentTarget
                    {
                        ActNumber = Article?.PublicationNumber,
                        ActYear = Article?.PublicationYear,
                        Article = articleNumber
                    });
                }

                AmendmentOperations.Add(amendmentOperation);
            }
        }
    }

    public class AmendmentOperation
    {
        public AmendmentOperationType Type { get; set; }
        public List<AmendmentTarget> AmendmentTargets { get; set; }
        public AmendmentOperation()
        {
            AmendmentTargets = new List<AmendmentTarget>();
        }
    }


    public class AmendmentTarget
    {
        public string? ActNumber { get; set; }
        public string? ActYear { get; set; }
        public string? Article { get; set; }
        public string? Subsection { get; set; }
        public string? Point { get; set; }
        public string? Letter { get; set; }
    }

    public enum AmendmentOperationType
    {
        [EnumDescription("uchylenie")]
        Repealed
    }
}