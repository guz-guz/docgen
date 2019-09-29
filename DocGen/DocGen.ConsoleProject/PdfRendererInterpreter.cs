using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Blip = DocumentFormat.OpenXml.Drawing.Blip;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;

namespace Test
{
    public class PdfRendererInterpreter : BaseInterpreter
    {
        private readonly IRendererVisitor _visitor;

        public PdfRendererInterpreter(InterpreterContext context, IRendererVisitor visitor)
            : base(context)
        {
            _visitor = visitor;
        }

        protected override bool VisitElement(InterpreterContext context, OpenXmlElement element)
        {
            switch (element.LocalName)
            {
                case Constants.SectionXmlProperties:
                    _visitor.VisitSection((SectionProperties) element);
                    break;

                case Constants.ParagraphXmlName:
                    _visitor.VisitParagraph((Paragraph) element);
                    break;

                case Constants.RunXmlName:
                    _visitor.VisitRun((Run) element);
                    break;

                case Constants.HeaderReferenceElement:
                    _visitor.VisitHeaderReference((HeaderReference) element);
                    break;

                case Constants.FooterReferenceElement:
                    _visitor.VisitFooterReference((FooterReference) element);
                    break;

                case Constants.BlipElement:
                    _visitor.VisitBlip((Blip) element);
                    break;
                
                case Constants.RunPropertiesXmlName:
                    if (element is RunProperties runProperties)
                    {
                        _visitor.VisitRunProperties(runProperties);
                    }
                    break;
                
                case Constants.TabsXmlName:
                    return _visitor.VisitTabs((Tabs) element);
                
                case Constants.TabXmlName:
                    _visitor.VisitTabSymbol();
                    break;
            }

            return true;
        }
    }
}