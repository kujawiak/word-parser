using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
    // Tytuł
    public class Title : BaseEntity 
    {   
        public string TitleText { get; set; }
        public List<Part> Parts { get; set; } = new List<Part>();

        public Title(Paragraph paragraph) : base(paragraph)
        {
        }
    }
}