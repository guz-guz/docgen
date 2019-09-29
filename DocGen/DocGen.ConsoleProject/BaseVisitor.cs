using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    public abstract class BaseVisitor
    {
        private const char FieldStartCharacter = '{';
        private const char FieldEndCharacter = '}';
        private const char FieldDelimiter = '.';

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

        protected string GetTextFragment(string runText)
        {
            var isField = false;
            List<string> chunks = null;
            var lastChunkStart = 0;
            var lastChunkEnd = 0;
            for (var i = 0; i < runText.Length; i++)
            {
                if (runText[i] == FieldStartCharacter)
                {
                    isField = true;
                    lastChunkStart = i;
                }
                else if (isField && runText[i] == FieldEndCharacter)
                {
                    chunks = chunks ?? new List<string>();
                    chunks.Add(runText.Substring(lastChunkEnd, lastChunkStart - lastChunkEnd));
                    var fieldText = runText.Substring(lastChunkStart + 1, i - lastChunkStart - 1);
                    chunks.Add(GetFieldValue(fieldText));
                    lastChunkEnd = i + 1;
                    isField = false;
                }
                else if (isField && !char.IsLetterOrDigit(runText[i]) && runText[i] != FieldDelimiter)
                {
                    isField = false;
                }
            }

            if (chunks == null)
            {
                return runText;
            }
            
            chunks.Add(runText.Substring(lastChunkEnd, runText.Length - lastChunkEnd));
            return string.Concat(chunks);
        }

        private string GetFieldValue(string fieldPath)
        {
            var pathPrefix = GetFieldPrefix(fieldPath);
            if (_pathsToValueProviders.TryGetValue(pathPrefix ?? string.Empty, out var valueProvider))
            {
                return Convert.ToString(valueProvider.GetValue(fieldPath));
            }

            var provider = _pathsToValueProviders.OrderByDescending(p => p.Key.Length)
                               .Where(p => pathPrefix.StartsWith(p.Key)).Select(p => p.Value).FirstOrDefault()
                           ?? throw new KeyNotFoundException($"Cannot find value provider for path {fieldPath}");
            var collection = (provider.GetValue(pathPrefix) as IEnumerable<object>)?.ToList()
                             ?? throw new InvalidOperationException($"Value at path {fieldPath} is not a collection");
            var tableRow = Context.ReturnToParent<TableRow>();
            PrepareDynamicRows(tableRow, pathPrefix, provider, collection);

            return fieldPath;
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