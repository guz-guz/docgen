using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using PdfDocument = MigraDoc.DocumentObjectModel.Document;
using PdfSection = MigraDoc.DocumentObjectModel.Section;
using PdfParagraph = MigraDoc.DocumentObjectModel.Paragraph;
using PdfTable = MigraDoc.DocumentObjectModel.Tables.Table;
using PdfTableRow = MigraDoc.DocumentObjectModel.Tables.Row;
using PdfTableCell = MigraDoc.DocumentObjectModel.Tables.Cell;
using PdfColor = MigraDoc.DocumentObjectModel.Color;
using PdfText = MigraDoc.DocumentObjectModel.FormattedText;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TabStop = DocumentFormat.OpenXml.Wordprocessing.TabStop;
using Underline = MigraDoc.DocumentObjectModel.Underline;

namespace Test
{
    public class PdfRendererVisitor : BaseVisitor, IRendererVisitor
    {
        private readonly PdfDocumentContext _documentContext;
        private readonly PdfDocument _document;
        private PdfSection _pdfSection;
        private PdfTable _pdfTable;
        private PdfTableRow _pdfRow;
        private int _currentCellIndex;
        private PdfTableCell _pdfCell;
        private PdfParagraph _pdfParagraph;
        private PdfText _pdfText;

        public PdfRendererVisitor(InterpreterContext context, PdfDocumentContext documentContext, object model = null)
            : base(context, model)
        {
            _documentContext = documentContext;
            _document = new PdfDocument();
        }

        public PdfDocument Document => _document;

        public void VisitRun(Run element)
        {
            var runText = element.InnerText;
            _pdfText = _pdfParagraph.AddFormattedText(GetTextFragment(runText));
        }

        public void VisitParagraph(Paragraph element)
        {
            if (_pdfSection == null)
            {
                _pdfSection = _document.AddSection();
            }

            _pdfParagraph = _pdfCell?.AddParagraph() ?? _pdfSection.AddParagraph();
            SetParagraphProperties(element.ParagraphProperties);
        }

        private void SetParagraphProperties(ParagraphProperties paragraphProperties)
        {
            if (paragraphProperties == null)
            {
                return;
            }
            
            if (paragraphProperties.ParagraphStyleId != null)
            {
                SetParagraphByStyles(_documentContext.Styles[paragraphProperties.ParagraphStyleId.Val]);
            }

            if (paragraphProperties.Indentation != null)
            {
                _pdfParagraph.Format.LeftIndent = GetPdfUnitLength(Convert.ToInt32(paragraphProperties.Indentation.Left.Value));
            }
        }

        private void SetParagraphByStyles(Style style)
        {
            if (style.StyleParagraphProperties.Indentation != null)
            {
                _pdfParagraph.Format.LeftIndent = GetPdfUnitLength(Convert.ToInt32(style.StyleParagraphProperties.Indentation.Left.Value));
            }
            
            SetParagraphByRunStyles(style.StyleRunProperties);
        }

        private void SetParagraphByRunStyles(StyleRunProperties element)
        {
            if (element == null)
            {
                return;
            }
            
            _pdfParagraph.Format.Font.Bold = element.Bold != null;
            _pdfParagraph.Format.Font.Italic = element.Italic != null;
            _pdfParagraph.Format.Font.Underline = element.Underline != null ? Underline.Single : Underline.None;
            if (element.FontSize != null)
            {
                _pdfParagraph.Format.Font.Size = Convert.ToSingle(element.FontSize.Val) / 2;
                _pdfParagraph.Format.Font.Name = element.RunFonts.Ascii?.Value ?? string.Empty;
                _pdfParagraph.Format.Font.Color = element.Color != null? PdfColor.Parse("#" + element.Color.Val): PdfColor.Empty;
            }
        }

        public void VisitRunProperties(RunProperties element)
        {
            if (_pdfParagraph == null)
            {
                return;
            }
            
            _pdfText.Font.Bold = element.Bold != null;
            _pdfText.Font.Italic = element.Italic != null;
            _pdfText.Font.Underline = element.Underline != null ? Underline.Single : Underline.None;
            if (element.FontSize != null)
            {
                _pdfText.Font.Size = Convert.ToSingle(element.FontSize.Val) / 2;    
                _pdfText.Font.Name = element.RunFonts.Ascii.Value;
                _pdfText.Font.Color = element.Color != null? PdfColor.Parse("#" + element.Color.Val): PdfColor.Empty;
            }
        }

        public bool VisitTabs(Tabs element)
        {
            foreach (TabStop tab in element)
            {
                if (tab.Val.Value == TabStopValues.Center)
                {
                    _pdfParagraph.Format.AddTabStop(GetPdfUnitLength(tab.Position), TabAlignment.Center);
                }
                else if (tab.Val.Value == TabStopValues.Left)
                {
                    _pdfParagraph.Format.AddTabStop(GetPdfUnitLength(tab.Position), TabAlignment.Left);
                }
                else if (tab.Val.Value == TabStopValues.Right)
                {
                    _pdfParagraph.Format.AddTabStop(GetPdfUnitLength(tab.Position), TabAlignment.Right);
                }
            }
            return false;
        }
        
        public void VisitTabSymbol()
        {
            _pdfParagraph.AddTab();
        }

        public void VisitSection(SectionProperties element)
        {
            var pageSize = element.GetFirstChild<PageSize>();
            _pdfSection.PageSetup.PageWidth = GetPdfUnitLength(pageSize.Width);
            _pdfSection.PageSetup.PageHeight = GetPdfUnitLength(pageSize.Height);
            var pageMargin = element.GetFirstChild<PageMargin>();
            _pdfSection.PageSetup.LeftMargin = GetPdfUnitLength(pageMargin.Left);
            _pdfSection.PageSetup.RightMargin = GetPdfUnitLength(pageMargin.Right);
            _pdfSection.PageSetup.TopMargin = _documentContext.TopMargin ?? GetPdfUnitLength(pageMargin.Top);
            _pdfSection.PageSetup.BottomMargin = _documentContext.BottomMargin ?? GetPdfUnitLength(pageMargin.Bottom);
            _pdfSection.PageSetup.HeaderDistance = GetPdfUnitLength(pageMargin.Header);
            _pdfSection.PageSetup.FooterDistance = GetPdfUnitLength(pageMargin.Footer);
        }

        public void VisitTable(Table table)
        {
            _pdfTable = _pdfCell == null
                ? _pdfSection.AddTable()
                : _pdfCell.Elements.AddTable();
        }

        public void VisitTableRow(TableRow element)
        {
            _pdfRow = _pdfTable.AddRow();
            if (TableRowsToDataProviders.TryGetValue(element, out var tableRowValueProvider))
            {
                tableRowValueProvider.MoveNext();
            }
        }

        public void VisitTableCell(TableCell tableCell)
        {
            _pdfCell = _pdfRow.Cells[_currentCellIndex++];
        }

        public void VisitHeaderReference(HeaderReference headerReference)
        {
            SetHeadersFooters(_pdfSection.Headers, headerReference);
        }

        public void VisitFooterReference(FooterReference headerReference)
        {
            SetHeadersFooters(_pdfSection.Footers, headerReference);
        }

        public void VisitBlip(Blip element)
        {
            var image = _documentContext.Images[element.Embed.Value];
            _pdfParagraph.AddImage($"base64:{Convert.ToBase64String(image)}");
        }

        public byte[] ToPdf()
        {
            var documentRenderer = new PdfDocumentRenderer(true) {Document = _document};
            documentRenderer.RenderDocument();
            using (var documentStream = new MemoryStream())
            {
                documentRenderer.PdfDocument.Save(documentStream);
                return documentStream.ToArray();
            }
        }

        protected override void PrepareDynamicRows(TableRow tableRow, string pathPrefix, ValueProvider provider,
            List<object> collection)
        {
            Context.PushElements(Enumerable.Repeat(tableRow, collection.Count));
            TableRowsToDataProviders.Add(tableRow,
                new TableRowValueProvider(collection.GetType().GetGenericArguments().First(), pathPrefix, collection)
            );
        }

        private int GetPdfUnitLength(float openXmlUnitLength)
        {
            return (int) (openXmlUnitLength * 72f / 1440f);
        }

        private void SetHeadersFooters(HeadersFooters headersFooters, HeaderFooterReferenceType referenceType)
        {
            switch (referenceType.Type.Value)
            {
                case HeaderFooterValues.Default:
                    CopyChildElements(headersFooters.Primary.Elements, referenceType.Id.Value);
                    break;

                case HeaderFooterValues.First:
                    CopyChildElements(headersFooters.FirstPage.Elements, referenceType.Id.Value);
                    break;
                
                case HeaderFooterValues.Even:
                    CopyChildElements(headersFooters.EvenPage.Elements, referenceType.Id.Value);
                    break;
            }
        }
        
        private void CopyChildElements(DocumentElements copyTo, string referenceId)
        {
            var relatedDocumentElements = _documentContext.RelatedDocuments[referenceId]
                .Sections
                .OfType<PdfSection>()
                .First()
                .Clone()    
                .Elements;

            foreach (DocumentObject element in relatedDocumentElements)
            {
                copyTo.Add(element.Clone() as DocumentObject);
            }
        }
    }
}