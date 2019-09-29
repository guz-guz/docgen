using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    class InPlaceRendererInterpreter: BaseInterpreter 
    {
        private readonly ITemplateVisitor _visitor;

        public InPlaceRendererInterpreter(InterpreterContext context, ITemplateVisitor visitor) 
            : base(context)
        {
            _visitor = visitor;
        }

        protected override void VisitElement(InterpreterContext context, OpenXmlElement element)
        {
            switch (element.XmlQualifiedName.ToString())
            {
                case Constants.RunXmlName:
                    _visitor.VisitRun((Run)element);
                    return;
                
                case Constants.TableRowXmlName: 
                    _visitor.VisitTableRow((TableRow)element);
                    return;
            }
        }
    }
}