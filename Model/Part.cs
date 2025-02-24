using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordParser.Model
{
    // Dział 
    public class Part 
    {
        Title Parent { get; set; }
        public string Number { get; set; }
        public List<Chapter> Chapters { get; set; }
    }
}