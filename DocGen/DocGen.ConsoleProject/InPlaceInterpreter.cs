using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    class InPlaceInterpreter: BaseInterpreter 
    {
        public ITemplateVisitor TemplateVisitor { get; set; }
        
        protected override void VisitElement(InterpreterContext context, OpenXmlElement element)
        {
            switch (element.XmlQualifiedName.ToString())
            {
                case Constants.RunXmlName:
                    TemplateVisitor.VisitRun(context, (Run)element);
                    return;
                
                case Constants.TableRowXmlName: 
                    TemplateVisitor.VisitTableRow(context, (TableRow)element);
                    return;
            }
        }
    }
}