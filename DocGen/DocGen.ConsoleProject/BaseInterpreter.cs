using DocumentFormat.OpenXml;

namespace Test
{
    public abstract class BaseInterpreter
    {
        private readonly InterpreterContext _context;

        protected BaseInterpreter(InterpreterContext context)
        {
            _context = context;
        }
        
        protected abstract void VisitElement(InterpreterContext context, OpenXmlElement element);

        public void Interpret(OpenXmlElement rootElement)
        {
            _context.PushElement(rootElement);
            while (_context.TryPopElement(out var element))
            {
                VisitElement(_context, element);
                _context.SetVisited(element);
                _context.PushReversedElements(element.ChildElements);
            }
        }
    }
}