using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    public abstract class BaseVisitor
    {
        private readonly Dictionary<string, ValueProvider> _pathsToValueProviders;

        protected readonly InterpreterContext Context;
        protected readonly Dictionary<TableRow, TableRowValueProvider> TableRowsToDataProviders;

        protected BaseVisitor(InterpreterContext context, object model = null)
        {
            _pathsToValueProviders = new Dictionary<string, ValueProvider>();
            if (model != null)
            {
                _pathsToValueProviders.Add(string.Empty, new ValueProvider(model.GetType()) {Model = model});
            }

            Context = context;
            TableRowsToDataProviders = new Dictionary<TableRow, TableRowValueProvider>();
        }

        protected abstract void PrepareDynamicRows(TableRow tableRow, string pathPrefix, ValueProvider provider,
            List<object> collection);

        protected bool TryGetParagraphText(string runText, out string text)
        {
            if (!GetIsField(runText))
            {
                text = runText;
                return true;
            }

            var fieldPath = runText.Trim('{', '}');
            var pathPrefix = GetFieldPrefix(fieldPath);
            if (_pathsToValueProviders.TryGetValue(fieldPath, out var valueProvider))
            {
                text = Convert.ToString(valueProvider.GetValue(fieldPath));
                return true;
            }

            var provider = _pathsToValueProviders.OrderByDescending(p => p.Key.Length)
                               .Where(p => pathPrefix.StartsWith(p.Key)).Select(p => p.Value).FirstOrDefault()
                           ?? throw new KeyNotFoundException($"Cannot find value provider for path {fieldPath}");
            var collection = (provider.GetValue(pathPrefix) as IEnumerable<object>)?.ToList()
                             ?? throw new InvalidOperationException($"Value at path {fieldPath} is not a collection");
            var tableRow = Context.ReturnToParent<TableRow>();
            PrepareDynamicRows(tableRow, pathPrefix, provider, collection);
            text = null;

            return false;
        }

        private string GetFieldPrefix(string path)
        {
            var delimiterIndex = path.LastIndexOf('.');
            return delimiterIndex >= 0
                ? path.Substring(0, delimiterIndex)
                : null;
        }

        private bool GetIsField(string text)
        {
            return text.StartsWith('{') && text.EndsWith('}');
        }
    }
}