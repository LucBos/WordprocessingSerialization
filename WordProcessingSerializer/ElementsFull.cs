using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordProcessingSerializer
{
    public class ElementsFull
    {
        public IEnumerable<OpenXmlElement> Elements { get; set; }
        public Numbering Numbering { get; set; }
        public IEnumerable<Style> Styles { get; set; }
    }
}