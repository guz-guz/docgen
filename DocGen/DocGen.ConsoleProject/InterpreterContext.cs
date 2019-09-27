using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;

namespace Test
{
    public class InterpreterContext
    {
        private Stack<OpenXmlElement> _elements = new Stack<OpenXmlElement>();
        private readonly Dictionary<int, OpenXmlElement> _distancesToElements = new Dictionary<int, OpenXmlElement>();

        public OpenXmlElement ReturnToParent<T>() where T : OpenXmlElement
        {
            var (key, value) = _distancesToElements
                .Where(p => p.Value is T)
                .OrderBy(p => p.Key)
                .First();
            _elements = new Stack<OpenXmlElement>(_elements.Take(key));
            return value;
        }

        public void PushElement(OpenXmlElement element)
        {
            _elements.Push(element);
        }

        public void PushElements(IEnumerable<OpenXmlElement> elements)
        {
            foreach (var childElement in elements)
            {
                _elements.Push(childElement);
            }
        }
            
        public bool TryPopElement(out OpenXmlElement element)
        {
            return _elements.TryPop(out element);
        }
            
        public void SetVisited(OpenXmlElement element)
        {
            _distancesToElements[_elements.Count] = element;
        }
    }
}