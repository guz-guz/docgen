using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace Test
{
    public interface IRendererVisitor: ITemplateVisitor
    {
        void VisitSection(SectionProperties element);
        void VisitParagraph(Paragraph element);
        void VisitRunProperties(RunProperties element);
        void VisitTable(Table table);
        void VisitTableCell(TableCell tableCell);
        void VisitHeaderReference(HeaderReference headerReference);
        void VisitFooterReference(FooterReference headerReference);
        void VisitBlip(Blip element);
    }
}