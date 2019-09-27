using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Header = DocumentFormat.OpenXml.Wordprocessing.Header;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using PdfDocument = MigraDoc.DocumentObjectModel.Document;
using PdfSection = MigraDoc.DocumentObjectModel.Section;
using PdfParagraph = MigraDoc.DocumentObjectModel.Paragraph;
using PdfTable = MigraDoc.DocumentObjectModel.Tables.Table;
using PdfTableRow = MigraDoc.DocumentObjectModel.Tables.Row;
using PdfTableCell = MigraDoc.DocumentObjectModel.Tables.Cell;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace Test
{
    public class PdfRendererVisitor : IRendererVisitor
    {
        private readonly Dictionary<string, ValueProvider> _pathsToValueProviders;
        private readonly Dictionary<Table, ValueProvider> _tablesToDataProviders;
        private bool _isDynamicTable;
        private PdfDocument _document;
        private PdfSection _pdfSection;
        private PdfTable _pdfTable;
        private PdfTableRow _pdfRow;
        private int _currentCellIndex;
        private PdfTableCell _pdfCell;
        private PdfParagraph _pdfParagraph;

        public PdfRendererVisitor()
        {
            _document = new PdfDocument();
            _pathsToValueProviders = new Dictionary<string, ValueProvider>();
            _tablesToDataProviders = new Dictionary<Table, ValueProvider>();
        }

        public void VisitRun(InterpreterContext context, Run element)
        {
            var runText = element.InnerText;

            if (TryGetParagraphText(context, runText, out string paragraphText))
            {
                _pdfParagraph.AddText(paragraphText);
            }
        }

        public void VisitTable(Table table)
        {
            _pdfTable = _pdfCell == null
                ? _pdfSection.AddTable()
                : _pdfCell.Elements.AddTable();
            //var properties = table.GetFirstChild<TableProperties>();
        }

        public void VisitTableRow(InterpreterContext context, TableRow element)
        {
            _pdfRow = _pdfTable.AddRow();
        }

        public void VisitTableCell(TableCell tableCell)
        {
            _pdfCell = _pdfRow.Cells[_currentCellIndex++];
        }

        public void VisitParagraph(Paragraph element)
        {
            _pdfParagraph = _pdfCell?.AddParagraph() ?? _pdfSection.AddParagraph();
        }

        public void VisitHeader(Header header)
        {
        }

        public void VisitFooter(Header header)
        {
        }
        
        private bool TryGetParagraphText(InterpreterContext context, string runText, out string text)
        {
            if (!GetIsField(runText))
            {
                text = runText;
                return true;
            }

            var fieldPath = runText.Trim('[', ']');
            var pathPrefix = GetFieldPrefix(fieldPath);
            if (_pathsToValueProviders.TryGetValue(fieldPath, out var valueProvider))
            {
                text = Convert.ToString(valueProvider.GetValue(fieldPath));
                return true;
            }
            
            var provider = _pathsToValueProviders.OrderByDescending(p => p.Key.Length).Where(p => pathPrefix.StartsWith(p.Key)).Select(p => p.Value).FirstOrDefault()
                ?? throw new KeyNotFoundException($"Cannot find value provider for path {fieldPath}");
            var collection = (provider.GetValue(pathPrefix) as IEnumerable<object>)?.ToList()
                ?? throw new InvalidOperationException($"Value at path {fieldPath} is not a collection");
            var tableRow = context.ReturnToParent<TableRow>();
            context.PushElements(Enumerable.Repeat(tableRow, collection.Count));
            text = null;
            
            return false;
        }

        private string GetFieldPrefix(string path)
        {
            var delimiterIndex = path.LastIndexOf('.');
            return delimiterIndex < 0
                ? path.Substring(0, delimiterIndex)
                : null;
        }

        private bool GetIsField(string text)
        {
            return text.StartsWith('[') && text.EndsWith(']');
        }
    }
}