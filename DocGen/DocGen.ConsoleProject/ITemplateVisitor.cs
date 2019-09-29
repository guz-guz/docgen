using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    public interface ITemplateVisitor
    {
        void VisitRun(Run element);
        void VisitTableRow(TableRow element);
    }
}