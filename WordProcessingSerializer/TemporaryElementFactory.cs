using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordProcessingSerializer
{
    internal class TemporaryElementFactory
    {
        public static OpenXmlElement GetElement(string outerXml)
        {
            if (outerXml.StartsWith("<w:p")) {
                return new Paragraph(outerXml);
            } 
            if (outerXml.StartsWith("<w:tbl")) {
                return new Table(outerXml);
            }
            throw new ArgumentException("Unknown element in the list");
        }
    }
}