using System.Collections.Generic;
using MigraDoc.DocumentObjectModel;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace Test
{
    public class PdfDocumentContext
    {
        public Dictionary<string, Document> RelatedDocuments { get; set; }
        public Dictionary<string, byte[]> Images { get; set; }
        public Dictionary<string, Style> Styles { get; set; }
        public int? TopMargin { get; set; }
        public int? BottomMargin { get; set; }
    }
}