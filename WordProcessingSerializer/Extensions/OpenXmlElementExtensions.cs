using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordProcessingSerializer
{
    static class OpenXmlElementExtensions
    {
        public static bool IsAnyOfType<T1>(this OpenXmlElement element)
        {
            return (element is T1) && element.Parent is Body;
        }

        public static bool IsAnyOfType<T1, T2>(this OpenXmlElement element)
        {
            return (element is T1 | element is T2) && element.Parent is Body;
        }

    }
}