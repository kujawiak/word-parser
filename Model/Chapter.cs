using System;
using System.Collections.Generic;

namespace WordParser.Model
{
    // Rozdział
    public class Chapter 
    {
        Part Parent { get; set; }
        public string Number { get; set; }
        public List<Section> Sections { get; set; }
    }
}