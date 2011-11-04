using System.IO;

namespace WordProcessingSerializer.Interfaces
{
    public interface IContentSerializer
    {
        void SerializeElementsBetweenBookmark(Stream stream, string copyFrom, string copyTo);
        void SerializeElementsFullBetweenBookmarks(Stream stream, string copyFrom, string copyTo);
    }
}