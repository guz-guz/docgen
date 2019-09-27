using DocumentFormat.OpenXml.Packaging;
using Document = MigraDoc.DocumentObjectModel.Document;

namespace Test
{
    public class Program
    {
        public static void Main()
        {
            using (var inputDocument = WordprocessingDocument.Open("Template.docx", true))
            {
                var outputDocument = new Document();
            }
        }
    }
}