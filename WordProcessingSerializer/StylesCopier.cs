using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordProcessingSerializer
{
    internal class StylesCopier
    {
        private readonly string _toDoc;
        private readonly string _fromDoc;

        public StylesCopier(string fromDoc, string toDoc)
        {
            _toDoc = toDoc;
            _fromDoc = fromDoc;
        }

        public StylesCopier(string toDoc)
        {
            _toDoc = toDoc;
        }

        public IEnumerable<Style> GetStylesFromDocument()
        {
            CheckIfFromDocumentSpecified();

            var fromWordDoc = WordprocessingDocument.Open(_fromDoc, true, new OpenSettings {AutoSave = false});
            return RetrieveStyles(fromWordDoc);
        }

        private static IEnumerable<Style> RetrieveStyles(WordprocessingDocument fromWordDoc)
        {
            using ((fromWordDoc))
            {
                var stylesDefinitionPart = fromWordDoc.MainDocumentPart.GetPartsOfType<StyleDefinitionsPart>().FirstOrDefault();

                if (stylesDefinitionPart == null)
                {
                    return new List<Style>();
                }

                return stylesDefinitionPart.Styles.AsEnumerable().OfType<Style>();
            }
        }

        public void ReplaceStyles(IEnumerable<Style> styles)
        {
            CheckIfToDocumentSpecified();

            var toWordDoc = WordprocessingDocument.Open(_toDoc, true, new OpenSettings());
            using ((toWordDoc))
            {
                StyleDefinitionsPart stylesDefinitionPart = toWordDoc.MainDocumentPart.GetPartsOfType<StyleDefinitionsPart>().FirstOrDefault();
                SynchronizeStylesWithSource(stylesDefinitionPart, styles);
            }
        }

        private void SynchronizeStylesWithSource(StyleDefinitionsPart stylesDefinitionPart, IEnumerable<Style> styles)
        {
            var styleNames = stylesDefinitionPart.Styles.OfType<Style>().Select(x => x.StyleName.Val).ToList();
            foreach (var style in styles)
            {
                if (!styleNames.Contains(style.StyleName.Val))
                {
                    stylesDefinitionPart.Styles.Append(style.CloneNode(true));
                }
            }
        }
        
        private void CheckIfToDocumentSpecified()
        {
            if (_toDoc == null)
            {
                throw new ArgumentException("A destination document must be provided");
            }
        }

        private void CheckIfFromDocumentSpecified()
        {
            if (_fromDoc == null)
            {
                throw new ArgumentException("A source document must be provided");
            }
        }
    }
}