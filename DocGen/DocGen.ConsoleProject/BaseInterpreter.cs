using DocumentFormat.OpenXml;

namespace Test
{
    abstract class BaseInterpreter
    {
        protected abstract void VisitElement(InterpreterContext context, OpenXmlElement element);

        public void Interpret(OpenXmlElement rootElement)
        {
            var context = new InterpreterContext();
            context.PushElement(rootElement);
            while (context.TryPopElement(out var element))
            {
                VisitElement(context, element);
                context.SetVisited(element);
                context.PushElements(element.ChildElements);
            }
        }
    }
}