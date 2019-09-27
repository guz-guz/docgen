using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    interface IRendererVisitor: ITemplateVisitor
    {
        void VisitParagraph(Paragraph element);
        void VisitTable(Table table);
        void VisitTableCell(TableCell tableCell);
        void VisitHeader(Header header);
        void VisitFooter(Header header);
    }
}