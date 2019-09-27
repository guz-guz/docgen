using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    interface ITemplateVisitor
    {
        void VisitRun(InterpreterContext context, Run element);
        void VisitTableRow(InterpreterContext context, TableRow element);
    }
}