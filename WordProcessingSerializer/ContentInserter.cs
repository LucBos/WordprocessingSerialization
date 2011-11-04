using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordProcessingSerializer.Interfaces;

namespace WordProcessingSerializer
{
    /// <summary>
    /// The ContentInserter will insert content right below the 'Paste' bookmark
    /// </summary>
    /// <remarks></remarks>
    public class ContentInserter : IContentInserter
    {

        private readonly string _fileLocation;
        public ContentInserter(string fileLocation)
        {
            _fileLocation = fileLocation;
        }

        public void InsertElementsInDocument(IEnumerable<OpenXmlElement> elements, string pasteBookmark)
        {
            var wordDoc = WordprocessingDocument.Open(_fileLocation, true);
            using ((wordDoc)) {
                var rootElement = wordDoc.MainDocumentPart.RootElement;
                var bookmark = rootElement.Descendants<BookmarkStart>().FirstOrDefault(x => x.Name.Value.Equals(pasteBookmark));

                if (bookmark == null) {
                    throw new ArgumentException("There was no bookmark found with the name " + pasteBookmark);
                }

                CleanTextBetweenBookmark(bookmark);
                var elementsList = elements.ToList();
                for (int i = elementsList.Count - 1; i >= 0; i--) {
                    bookmark.Parent.InsertAfterSelf(elementsList[i]);
                }
                bookmark.Parent.Remove();
            }
        }

        private void CleanTextBetweenBookmark(BookmarkStart bookmark)
        {
            var parentOfBookmark = bookmark.Parent;

            if (parentOfBookmark is Paragraph) {
                foreach (var text in parentOfBookmark.Descendants<Text>()) {
                    text.Remove();
                }


            } else if (parentOfBookmark is Body) {
                var bookmarkEnd = parentOfBookmark.Descendants<BookmarkEnd>().FirstOrDefault(x => x.Id.Equals(bookmark.Id));

                if (bookmarkEnd != null)
                {
                    var elementsToRemove = bookmark.ElementsAfter().Intersect(bookmarkEnd.ElementsBefore()).OfType<Text>();
                    foreach (var text in elementsToRemove)
                    {
                        text.Remove();
                    }
                }
            }
        }

        public void InsertElementsWithNumberingInDocument(ElementsFull elementsFull,string pasteBookmark)
        {
            InsertElementsInDocument(elementsFull.Elements,pasteBookmark);

            var numberingCopier = new NumberingCopier(_fileLocation);
            numberingCopier.ReplaceNumbering(elementsFull.Numbering);

            var stylesCopier = new StylesCopier(_fileLocation);
            stylesCopier.ReplaceStyles(elementsFull.Styles);
        }

    }
}