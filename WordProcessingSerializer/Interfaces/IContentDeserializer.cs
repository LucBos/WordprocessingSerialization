using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;

namespace WordProcessingSerializer.Interfaces
{
    public interface IContentDeserializer
    {
        IEnumerable<OpenXmlElement> DeserializeContent(Stream stream);
        ElementsFull DeserializeContentWithNumberingAndStyles(Stream stream);
    }
}