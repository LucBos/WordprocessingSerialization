using System.Collections.Generic;
using DocumentFormat.OpenXml;

namespace WordProcessingSerializer.Interfaces
{
    public interface IContentInserter
    {
        void InsertElementsInDocument(IEnumerable<OpenXmlElement> elements, string pasteBookmark);
        void InsertElementsWithNumberingInDocument(ElementsFull elementsFull, string pasteBookmark);
    }
}