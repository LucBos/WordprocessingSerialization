using System.Linq;
using WordProcessingSerializer.Interfaces;

namespace WordProcessingSerializer
{
    using System.Collections.Generic;
    using System.IO;
    using System.Runtime.Serialization.Formatters.Binary;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// The ContentSerializer will serialize the text between the CopyFrom bookmark en the CopyTo bookmark to a stream.
    /// </summary>
    /// <remarks></remarks>
    public class ContentSerializer : IContentSerializer
    {

        private readonly DocxReader _docxReader;
        public ContentSerializer(string fileLocation)
            : this(new DocxReader(fileLocation))
        {
        }

        public ContentSerializer(DocxReader docxReader)
        {
            _docxReader = docxReader;
        }

        public void SerializeElementsBetweenBookmark(Stream stream, string copyFrom, string copyTo)
        {
            var elementsBetweenBookmarks = _docxReader.GetElementsBetweenBookmarksFromFile(copyFrom, copyTo);
            WriteElementsToStream(elementsBetweenBookmarks, stream);
        }

        public void SerializeElementsFullBetweenBookmarks(Stream stream, string copyFrom, string copyTo)
        {
            var elementsBetweenBookmarks = _docxReader.GetElementsBetweenBookmarksFromFile(copyFrom, copyTo);
            var numberingPart = _docxReader.GetNumberingPart();
            var styles = _docxReader.GetStyles();
            WriteElementsFullToStream(elementsBetweenBookmarks, numberingPart, styles, stream);
        }

        private void WriteElementsToStream(IEnumerable<OpenXmlElement> elementsFromBookmarks, Stream stream)
        {
            // the IEnumerable of openXmlElements must be converted to an IEnumerable of strings before it can be converted
            var serializableIEnumerable = elementsFromBookmarks.Select(x => x.OuterXml).ToList();
            var binSerializer = new BinaryFormatter();
            binSerializer.Serialize(stream, serializableIEnumerable);
        }

        private void WriteElementsFullToStream(IEnumerable<OpenXmlElement> elementsFromBookmarks, Numbering numberingPart, IEnumerable<Style> styles, Stream stream)
        {
            // the IEnumerable of openXmlElements must be converted to an IEnumerable of strings before it can be converted
            var elementsWithNumbering = SerializableElementsFull.CreateFromElements(elementsFromBookmarks, numberingPart, styles);
            var binSerializer = new BinaryFormatter();
            binSerializer.Serialize(stream, elementsWithNumbering);
        }
    }
}
