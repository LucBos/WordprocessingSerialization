using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

namespace WordProcessingSerializer
{
    public class DocxReader
    {

        private readonly string _fileLocation;
        public DocxReader(string fileLocation)
        {
            if (fileLocation == null)
            {
                throw new ArgumentException("A filelocation must be specified!");
            }
            _fileLocation = fileLocation;

            AcceptRevisions();
        }

        private void AcceptRevisions()
        {
            var wordDoc = WordprocessingDocument.Open(_fileLocation, true);
            using (wordDoc)
            {
                RevisionAccepter.AcceptRevisions(wordDoc);
            }
        }

        public Numbering GetNumberingPart()
        {
            return new NumberingCopier(_fileLocation, null).GetFromNumberingPart();
        }

        public IEnumerable<OpenXmlElement> GetElementsBetweenBookmarksFromFile(string copyFromBookmark, string copyToBookmark)
        {
            var wordDoc = WordprocessingDocument.Open(_fileLocation, true);
            IEnumerable<OpenXmlElement> elementsBetweenBookmarks;
            using ((wordDoc))
            {
                elementsBetweenBookmarks = GetAllElementsBetweenBookmarks(wordDoc, copyFromBookmark, copyToBookmark);
            }
            return elementsBetweenBookmarks;
        }

        private IEnumerable<OpenXmlElement> GetAllElementsBetweenBookmarks(WordprocessingDocument wordDoc, object copyFrom, object copyTo)
        {

            var rootElement = wordDoc.MainDocumentPart.RootElement;
            var allElements = GetDescendatsFilteredByType(rootElement);

            var allBookmarks = GetAllBookmarks(rootElement).ToList();

            BookmarkStart copyToBookmark = GetCopyToBookmark(copyTo, allBookmarks);
            var copyFromBookmark = GetCopyFromBookmark(copyFrom, allBookmarks);

            OpenXmlElement fromParent = copyFromBookmark.GetElementAfter();
            OpenXmlElement toParent = copyToBookmark.GetElementBefore();

            var fromIndex = allElements.IndexOf(fromParent);
            var toIndex = allElements.IndexOf(toParent);

            return allElements.GetRange(fromIndex, toIndex - fromIndex + 1).Select(x => x.CloneNode(true)).ToList();
        }

        private static BookmarkStart GetCopyFromBookmark(object copyFrom, List<BookmarkStart> allBookmarks)
        {
            var copyFromBookmark = allBookmarks.FirstOrDefault(x => x.Name.InnerText.Equals(copyFrom));

            if (copyFromBookmark == null)
            {
                throw new InvalidOperationException(
                    "The content serializer can only serialize content between the copyFrom and the copyTo bookmarks.  Please make sure these are defined in the document!");
            }
            return copyFromBookmark;
        }

        private static BookmarkStart GetCopyToBookmark(object copyTo, List<BookmarkStart> allBookmarks)
        {
            var copyToBookmark = allBookmarks.FirstOrDefault(x => x.Name.InnerText.Equals(copyTo));

            if (copyToBookmark == null)
            {
                throw new InvalidOperationException(
                    "The content serializer can only serialize content between the copyFrom and the copyTo bookmarks.  Please make sure these are defined in the document!");
            }
            return copyToBookmark;
        }

        private static IEnumerable<BookmarkStart> GetAllBookmarks(OpenXmlPartRootElement rootElement)
        {
            var allBookmarks = rootElement.Descendants<BookmarkStart>();

            if (allBookmarks == null)
            {
                throw new InvalidOperationException("There were no bookmarks found in the source document.");
            }
            return allBookmarks;
        }

        private List<OpenXmlElement> GetDescendatsFilteredByType(OpenXmlPartRootElement rootElement)
        {
            return rootElement.Descendants().Where(x => x.IsAnyOfType<Paragraph, Table>()).ToList();
        }

        public IEnumerable<Style> GetStyles()
        {
            return new StylesCopier(_fileLocation, null).GetStylesFromDocument();
        }
    }
}