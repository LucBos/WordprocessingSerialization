using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordProcessingSerializer
{
    /// <summary>
    /// This class will copy the numbering styles from the fromDocument to the toDocument
    /// </summary>
    /// <remarks></remarks>
    internal class NumberingCopier
    {
        private readonly string _fromDoc;
        private readonly string _toDoc;

        public NumberingCopier(string fromDoc, string toDoc)
        {
            _toDoc = toDoc;
            _fromDoc = fromDoc;
        }

        public NumberingCopier(string toDoc)
        {
            _toDoc = toDoc;
        }

        public void ReplaceNumbering()
        {
            CheckIfFromDocumentSpecified();

            var fromWordDoc = WordprocessingDocument.Open(_fromDoc, true, new OpenSettings { AutoSave = false });
            using ((fromWordDoc))
            {
                var fromNumberingPart = fromWordDoc.MainDocumentPart.GetPartsOfType<NumberingDefinitionsPart>().FirstOrDefault();
                if (fromNumberingPart != null)
                    ReplaceNumbering(fromNumberingPart.Numbering);
            }
        }

        public Numbering GetFromNumberingPart()
        {
            CheckIfFromDocumentSpecified();

            var fromWordDoc = WordprocessingDocument.Open(_fromDoc, true, new OpenSettings { AutoSave = false });
            using ((fromWordDoc))
            {
                var fromNumberingPart = fromWordDoc.MainDocumentPart.GetPartsOfType<NumberingDefinitionsPart>().FirstOrDefault();
                return fromNumberingPart.Numbering;
            }
        }

        public void ReplaceNumbering(Numbering fromNumbering)
        {
            CheckIfToDocumentSpecified();

            var toWordDoc = WordprocessingDocument.Open(_toDoc, true);
            using ((toWordDoc))
            {
                var toNumberingPart = toWordDoc.MainDocumentPart.GetPartsOfType<NumberingDefinitionsPart>().FirstOrDefault();

                SyncNumberingBetweenDocuments(fromNumbering, toNumberingPart);
            }
        }

        private void SyncNumberingBetweenDocuments(Numbering fromNumbering, NumberingDefinitionsPart toNumberingPart)
        {
            FixNumberingInstances(fromNumbering, toNumberingPart);
            FixAbstractEnums(fromNumbering, toNumberingPart);
        }

        private void FixAbstractEnums(Numbering fromNumbering, NumberingDefinitionsPart toNumberingPart)
        {
            foreach (AbstractNum abstractNum in fromNumbering.Descendants<AbstractNum>())
            {
                if (!toNumberingPart.Numbering.Descendants<AbstractNum>().Any(x => x.AbstractNumberId.Value.Equals(abstractNum.AbstractNumberId.Value)))
                {
                    toNumberingPart.Numbering.Append(abstractNum.CloneNode(true));
                }
            }
        }

        private void FixNumberingInstances(Numbering fromNumbering, NumberingDefinitionsPart toNumberingPart)
        {
            foreach (NumberingInstance numberingInstance in fromNumbering.Descendants<NumberingInstance>())
            {
                if (!toNumberingPart.Numbering.Descendants<NumberingInstance>().Any(x => x.NumberID.Value.Equals(numberingInstance.NumberID.Value)))
                {
                    toNumberingPart.Numbering.Append(numberingInstance.CloneNode(true));
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