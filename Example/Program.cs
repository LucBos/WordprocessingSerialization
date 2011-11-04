using System;
using System.IO;
using WordProcessingSerializer;
using WordProcessingSerializer.Interfaces;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            {
                // The stream to which we'll be serializing and deserializing
                var memoryStream = new MemoryStream();

                // The content serializer is responsible for serializing the content between two bookmarks
                IContentSerializer serializer = new ContentSerializer("c:\\temp\\Test.docx");
                serializer.SerializeElementsFullBetweenBookmarks(memoryStream,
                                                                "profielstart",
                                                                "profieleind");

                Console.WriteLine("Content serialized");

                // If you want to write to the database you can do
                var data = memoryStream.ToArray();

                // If you then want to insert data back into the document you can do like this
                memoryStream = new MemoryStream(data);

                // The content deserializer is responsible for deserializing the serialized OpenXml content
                IContentDeserializer deserializer = new ContentDeserializer();
                var contentWithNumbering = deserializer.DeserializeContentWithNumberingAndStyles(memoryStream);

                Console.WriteLine("Content deserialized");

                // The content inserter will insert Paragraph and table and will synchronize the styles and numbering
                IContentInserter contentInserter = new ContentInserter("c:\\temp\\InsertInDocument.docx");
                contentInserter.InsertElementsWithNumberingInDocument(contentWithNumbering, "Paste");

                Console.WriteLine("Content inserted");
                Console.Read();
            }

        }
    }
}
