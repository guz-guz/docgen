using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace Test
{
    public class Program
    {
        public static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (var inputDocument = WordprocessingDocument.Open("Simplest.docx", false))
            {
                var documentContext = new PdfDocumentContext
                {
                    TopMargin = 120,
                    BottomMargin = 100
                };
                SetStyles(documentContext, inputDocument.MainDocumentPart);
                SetHeaderFooterDocuments(documentContext, inputDocument.MainDocumentPart);
                var mainVisitor = InterpretDocument(inputDocument.MainDocumentPart, documentContext);
                File.WriteAllBytes("result.pdf", mainVisitor.ToPdf());
            }
        }

        private static void SetStyles(PdfDocumentContext documentContext, OpenXmlPart openXmlPart)
        {
            var stylesPart = openXmlPart.GetPartsOfType<StylesPart>().First();
            documentContext.Styles = stylesPart.Styles.ChildElements
                .OfType<Style>()
                .Where(s => s.Type.Value == StyleValues.Paragraph)
                .ToDictionary(s => s.StyleId.Value, s => s);
        }

        private static void SetHeaderFooterDocuments(PdfDocumentContext documentContext, OpenXmlPart openXmlPart)
        {
            documentContext.RelatedDocuments = openXmlPart.Parts
                .Where(pair => pair.OpenXmlPart.RelationshipType == Constants.HeaderPartRelationType
                               || pair.OpenXmlPart.RelationshipType == Constants.FooterPartRelationType)
                .ToDictionary(pair => pair.RelationshipId, pair => InterpretDocument(pair.OpenXmlPart, documentContext).Document);
        }
        
        private static Dictionary<string, byte[]> GetImages(OpenXmlPart openXmlPart)
        {
            return openXmlPart.Parts
                .Where(pair => pair.OpenXmlPart.RelationshipType == Constants.ImagePartRelationType)
                .ToDictionary(pair => pair.RelationshipId, pair => ReadStream(((ImagePart)pair.OpenXmlPart).GetStream()));
        }

        private static byte[] ReadStream(Stream stream)
        {
            var buffer = new byte[stream.Length];
            stream.Read(buffer);
            return buffer;
        }
        
        private static PdfRendererVisitor InterpretDocument(OpenXmlPart openXmlPart, PdfDocumentContext documentContext)
        {
            var interpreterContext = new InterpreterContext();
            var visitor = new PdfRendererVisitor(interpreterContext, documentContext);
            var pdfInterpreter = new PdfRendererInterpreter(interpreterContext, visitor);
            documentContext.Images = GetImages(openXmlPart);
            pdfInterpreter.Interpret(openXmlPart.RootElement);
            return visitor;
        }
    }
}