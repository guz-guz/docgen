using System;
using System.Collections.Generic;

namespace Test
{
    public class TableRowValueProvider: ValueProvider
    {
        private readonly IEnumerator<object> _modelEnumerator;

        public TableRowValueProvider(Type modelType, string prefix, IEnumerable<object> modelEnumerator) 
            : base(modelType, prefix)
        {
            _modelEnumerator = modelEnumerator.GetEnumerator();
        }

        public bool MoveNext()
        {
            var result = _modelEnumerator.MoveNext();
            Model = _modelEnumerator.Current;
            return result;
        }
    }
}