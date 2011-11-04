using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordProcessingSerializer
{
    [Serializable()]
    public class SerializableElementsFull
    {
        public IEnumerable<string> Elements { get; set; }
        public string Numbering { get; set; }
        public IEnumerable<string> Styles { get; set; }

        public static SerializableElementsFull CreateFromElements(IEnumerable<OpenXmlElement> elementsFromBookmarks,
                                                                               Numbering numberingPart, IEnumerable<Style> styles)
        {
            var serializableElementsEnumerable = elementsFromBookmarks.Select(x => x.OuterXml).ToList();
            var serializableNumbering = numberingPart.OuterXml;
            var serializableStylesEnumerable = styles.Select(x => x.OuterXml).ToList();
            var elementsWithNumbering = new SerializableElementsFull
                                            {
                                                Numbering = serializableNumbering,
                                                Elements = serializableElementsEnumerable,
                                                Styles = serializableStylesEnumerable
                                            };
            return elementsWithNumbering;
        }
    }
}