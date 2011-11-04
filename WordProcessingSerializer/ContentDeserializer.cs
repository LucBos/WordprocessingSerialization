using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordProcessingSerializer.Interfaces;

namespace WordProcessingSerializer
{
    /// <summary>
    /// The ContentDeserializer will deserialize a binary stream to an IEnumerable of OpenXmlElements
    /// </summary>
    /// <remarks></remarks>
    public class ContentDeserializer : IContentDeserializer
    {
        public IEnumerable<OpenXmlElement> DeserializeContent(Stream stream)
        {
            var binSerializer = new BinaryFormatter();
            var serializedValue = binSerializer.Deserialize(stream) as List<string>;

            if(serializedValue == null)
            {
                throw new ArgumentException("Invalid serialized format.");
            }

            return serializedValue.Select(TemporaryElementFactory.GetElement);
        }

        public ElementsFull DeserializeContentWithNumberingAndStyles(Stream stream)
        {
            var binSerializer = new BinaryFormatter();
            var serializedValue = binSerializer.Deserialize(stream) as SerializableElementsFull;

            if (serializedValue == null)
            {
                throw new ArgumentException("Invalid serialized format.");
            }

            var paragraphsWithNumbering = new ElementsFull {
                                                               Elements = serializedValue.Elements.Select(TemporaryElementFactory.GetElement),
                                                               Numbering = new Numbering(serializedValue.Numbering),
                                                               Styles = serializedValue.Styles.Select(x => new Style(x))
                                                           };
            return paragraphsWithNumbering;
        }
    }
}