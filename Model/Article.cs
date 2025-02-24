using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
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
}