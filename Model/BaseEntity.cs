using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

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
        public BaseEntity(Paragraph paragraph)
        {
            Paragraph = paragraph;
            Content = paragraph.InnerText.Sanitize();
        }
    }
}