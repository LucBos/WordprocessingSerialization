using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordProcessingSerializer
{
    static class BookmarkExtensions
    {
        public static OpenXmlElement GetElementBefore(this BookmarkStart copyToBookmark)
        {
            var toParent = copyToBookmark.Parent;
            if (toParent is Body) {
                toParent = copyToBookmark.ElementsBefore().LastOrDefault();
            }
            return toParent;
        }

        public static OpenXmlElement GetElementAfter(this BookmarkStart copyFromBookmark)
        {
            var fromParent = copyFromBookmark.Parent;
            if (fromParent is Body) {
                fromParent = copyFromBookmark.ElementsAfter().ElementAt(1);
            }
            return fromParent;
        }
    }
}