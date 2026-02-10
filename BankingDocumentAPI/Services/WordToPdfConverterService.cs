using BankingDocumentAPI.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Linq;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using LineSpacingRuleValues = DocumentFormat.OpenXml.Wordprocessing.LineSpacingRuleValues;

namespace BankingDocumentAPI.Services
{
    /// <summary>
    /// Service for converting Word documents (DOCX) to PDF using QuestPDF
    /// by parsing the OpenXML structure and recreating it in QuestPDF
    /// </summary>
    public class WordToPdfConverterService
    {
        private readonly ILogger<WordToPdfConverterService> _logger;

        public WordToPdfConverterService(ILogger<WordToPdfConverterService> logger)
        {
            _logger = logger;
            QuestPDF.Settings.License = LicenseType.Community;
        }

        /// <summary>
        /// Convert Word document bytes to PDF bytes
        /// </summary>
        public byte[] ConvertWordBytesToPdf(byte[] wordBytes, Dictionary<string, string>? placeholders = null)
        {
            try
            {
                // Parse the Word document
                var parsedDoc = ParseWordDocument(wordBytes);

                // If placeholders provided, replace them in parsed content
                if (placeholders != null && placeholders.Count > 0)
                {
                    ReplacePlaceholders(parsedDoc, placeholders);
                }

                // Generate PDF from parsed document
                return GeneratePdfFromParsed(parsedDoc);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error converting Word to PDF");
                throw;
            }
        }

        /// <summary>
        /// Parse Word document from bytes into structured format
        /// </summary>
        private ParsedWordDocument ParseWordDocument(byte[] wordBytes)
        {
            var parsedDoc = new ParsedWordDocument();

            using var memoryStream = new MemoryStream(wordBytes);
            using var wordDocument = WordprocessingDocument.Open(memoryStream, false);

            var mainPart = wordDocument.MainDocumentPart;
            if (mainPart == null)
                throw new InvalidOperationException("Invalid Word document: No MainDocumentPart");

            var body = mainPart.Document.Body;
            if (body == null)
                return parsedDoc;

            // Parse headers
            foreach (var headerPart in mainPart.HeaderParts)
            {
                if (parsedDoc.Header == null)
                {
                    parsedDoc.Header = ParseHeader(headerPart.Header);
                }
            }

            // Parse footers
            foreach (var footerPart in mainPart.FooterParts)
            {
                if (parsedDoc.Footer == null)
                {
                    parsedDoc.Footer = ParseHeader(footerPart.Footer);
                }
            }

            // Parse body elements (paragraphs and tables)
            foreach (var openXmlElement in body.ChildElements)
            {
                if (openXmlElement is Paragraph paragraph)
                {
                    var paraElement = ParseParagraph(paragraph);
                    parsedDoc.BodyElements.Add(new ParagraphDocumentElement { Paragraph = paraElement });
                }
                else if (openXmlElement is Table table)
                {
                    var tableElement = ParseTable(table);
                    parsedDoc.BodyElements.Add(new TableDocumentElement { Table = tableElement });
                }
            }

            return parsedDoc;
        }

        /// <summary>
        /// Parse header or footer from header/footer element
        /// </summary>
        private HeaderFooterElement ParseHeader(Header? header)
        {
            if (header == null) return new HeaderFooterElement();

            var element = new HeaderFooterElement();

            foreach (var paragraph in header.Elements<Paragraph>())
            {
                var paraElement = ParseParagraph(paragraph);
                element.Paragraphs.Add(paraElement);

                // Check for page numbers
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        // Simple check for page number patterns
                        if (text.Text?.Contains("PAGE", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            element.HasPageNumber = true;
                        }
                    }
                }
            }

            return element;
        }

        private HeaderFooterElement ParseHeader(Footer? footer)
        {
            if (footer == null) return new HeaderFooterElement();

            var element = new HeaderFooterElement();

            foreach (var paragraph in footer.Elements<Paragraph>())
            {
                var paraElement = ParseParagraph(paragraph);
                element.Paragraphs.Add(paraElement);

                // Check for page numbers
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        if (text.Text?.Contains("PAGE", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            element.HasPageNumber = true;
                        }
                    }
                }
            }

            return element;
        }

        /// <summary>
        /// Parse a Word paragraph into our ParagraphElement format
        /// </summary>
        private ParagraphElement ParseParagraph(Paragraph wordPara)
        {
            var element = new ParagraphElement();

            // Parse alignment
            var justification = wordPara.ParagraphProperties?.Justification;
            if (justification != null)
            {
                var val = justification.Val?.Value;
                if (val == JustificationValues.Center)
                    element.Alignment = Models.TextAlignment.Center;
                else if (val == JustificationValues.Right)
                    element.Alignment = Models.TextAlignment.Right;
                else if (val == JustificationValues.Both)
                    element.Alignment = Models.TextAlignment.Justify;
                else
                    element.Alignment = Models.TextAlignment.Left;
            }

            // Parse spacing
            var spacing = wordPara.ParagraphProperties?.SpacingBetweenLines;
            if (spacing != null)
            {
                element.Spacing = new SpacingInfo
                {
                    Before = ConvertToPoints(spacing.Before?.Value),
                    After = ConvertToPoints(spacing.After?.Value),
                    LineSpacing = ParseLineSpacing(spacing.Line?.Value, spacing.LineRule?.Value)
                };
            }

            // Parse indentation
            var indentation = wordPara.ParagraphProperties?.Indentation;
            if (indentation != null)
            {
                element.Indent = new IndentInfo
                {
                    Left = ConvertToPoints(indentation.Left?.Value),
                    Right = ConvertToPoints(indentation.Right?.Value),
                    FirstLine = ConvertToPoints(indentation.FirstLine?.Value)
                };
            }

            // Check if heading
            var style = wordPara.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (style != null && style.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            {
                element.IsHeading = true;
                if (style.Length > 7 && int.TryParse(style.Substring(7), out int level))
                {
                    element.HeadingLevel = level;
                }
            }

            // Parse runs
            foreach (var run in wordPara.Elements<Run>())
            {
                var textRun = ParseRun(run);
                element.Runs.Add(textRun);
            }

            return element;
        }

        /// <summary>
        /// Parse a Word run into our TextRun format
        /// </summary>
        private TextRun ParseRun(Run run)
        {
            var textRun = new TextRun();
            var props = run.RunProperties;

            if (props != null)
            {
                textRun.Bold = props.Bold != null;
                textRun.Italic = props.Italic != null;
                textRun.Underline = props.Underline != null;

                // Font size (half-points to points)
                if (props.FontSize != null && props.FontSize.Val != null)
                {
                    if (int.TryParse(props.FontSize.Val.Value, out int sizeHalfPoints))
                    {
                        textRun.FontSize = sizeHalfPoints / 2;
                    }
                }

                // Font family
                if (props.RunFonts != null)
                {
                    textRun.FontFamily = props.RunFonts.Ascii?.Value
                        ?? props.RunFonts.HighAnsi?.Value
                        ?? "Arial";
                }

                // Color
                if (props.Color != null && !string.IsNullOrEmpty(props.Color.Val?.Value))
                {
                    textRun.Color = "#" + props.Color.Val.Value;
                }
            }

            // Extract text
            var textContent = string.Concat(run.Elements<Text>().Select(t => t.Text ?? ""));
            textRun.Text = textContent;

            return textRun;
        }

        /// <summary>
        /// Parse a Word table into our TableElement format
        /// </summary>
        private TableElement ParseTable(Table wordTable)
        {
            var tableElement = new TableElement();
            var rows = wordTable.Elements<TableRow>().ToList();

            if (rows.Count == 0)
                return tableElement;

            // Determine column count from first row
            var firstRow = rows[0];
            tableElement.ColumnCount = firstRow.Elements<TableCell>().Count();

            // Parse each row
            foreach (var wordRow in rows)
            {
                var rowElement = new TableRowElement();
                var cells = wordRow.Elements<TableCell>().ToList();

                // Check if header row (typically first row with bold text)
                rowElement.IsHeader = wordRow == firstRow && IsHeaderRow(cells);

                foreach (var wordCell in cells)
                {
                    var cellElement = ParseTableCell(wordCell);
                    rowElement.Cells.Add(cellElement);
                }

                tableElement.Rows.Add(rowElement);
            }

            return tableElement;
        }

        /// <summary>
        /// Parse a Word table cell
        /// </summary>
        private TableCellElement ParseTableCell(TableCell wordCell)
        {
            var cellElement = new TableCellElement();

            var cellProps = wordCell.TableCellProperties;
            if (cellProps != null)
            {
                // Column span
                if (cellProps.GridSpan != null && cellProps.GridSpan.Val != null)
                {
                    // Try to parse the grid span value
                    try
                    {
                        var innerText = cellProps.GridSpan.Val.InnerText;
                        if (!string.IsNullOrEmpty(innerText))
                        {
                            cellElement.ColumnSpan = int.Parse(innerText);
                        }
                    }
                    catch
                    {
                        cellElement.ColumnSpan = 1;
                    }
                }

                // Note: Vertical alignment parsing is skipped due to OpenXML API complexity
                // Default to Top alignment

                // Background color
                if (cellProps.Shading != null && !string.IsNullOrEmpty(cellProps.Shading.Fill?.Value))
                {
                    cellElement.BackgroundColor = "#" + cellProps.Shading.Fill.Value;
                }

                // Width
                if (cellProps.TableCellWidth != null)
                {
                    cellElement.CellWidth = ConvertToPoints(cellProps.TableCellWidth.Width?.Value);
                }
            }

            // Parse paragraphs in cell
            foreach (var paragraph in wordCell.Elements<Paragraph>())
            {
                var paraElement = ParseParagraph(paragraph);
                cellElement.Paragraphs.Add(paraElement);
            }

            return cellElement;
        }

        /// <summary>
        /// Check if a row is a header row
        /// </summary>
        private bool IsHeaderRow(List<TableCell> cells)
        {
            foreach (var cell in cells)
            {
                foreach (var para in cell.Elements<Paragraph>())
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        if (run.RunProperties?.Bold != null)
                            return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Convert OpenXML spacing value to points
        /// </summary>
        private double ConvertToPoints(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return 0;

            // OpenXML uses twips (1/20 of a point) for some values
            if (int.TryParse(value, out int twips))
            {
                return twips / 20.0;
            }
            return 0;
        }

        /// <summary>
        /// Parse line spacing value
        /// </summary>
        private double ParseLineSpacing(string? value, LineSpacingRuleValues? rule)
        {
            if (string.IsNullOrEmpty(value))
                return 1.0;

            if (int.TryParse(value, out int lineSpacing))
            {
                if (rule == LineSpacingRuleValues.Auto)
                    return lineSpacing / 240.0; // Lines
                else if (rule == LineSpacingRuleValues.AtLeast)
                    return lineSpacing / 20.0; // Points
            }
            return 1.0;
        }

        /// <summary>
        /// Replace placeholders in parsed document
        /// </summary>
        private void ReplacePlaceholders(ParsedWordDocument doc, Dictionary<string, string> placeholders)
        {
            foreach (var element in doc.BodyElements)
            {
                if (element is ParagraphDocumentElement paraElement)
                {
                    ReplacePlaceholdersInParagraph(paraElement.Paragraph, placeholders);
                }
                else if (element is TableDocumentElement tableElement)
                {
                    ReplacePlaceholdersInTable(tableElement.Table, placeholders);
                }
            }

            if (doc.Header != null)
            {
                foreach (var para in doc.Header.Paragraphs)
                {
                    ReplacePlaceholdersInParagraph(para, placeholders);
                }
            }

            if (doc.Footer != null)
            {
                foreach (var para in doc.Footer.Paragraphs)
                {
                    ReplacePlaceholdersInParagraph(para, placeholders);
                }
            }
        }

        private void ReplacePlaceholdersInParagraph(ParagraphElement paragraph, Dictionary<string, string> placeholders)
        {
            foreach (var run in paragraph.Runs)
            {
                foreach (var placeholder in placeholders)
                {
                    run.Text = run.Text.Replace($"{{{placeholder.Key}}}", placeholder.Value);
                }
            }
        }

        private void ReplacePlaceholdersInTable(TableElement table, Dictionary<string, string> placeholders)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        ReplacePlaceholdersInParagraph(para, placeholders);
                    }
                }
            }
        }

        /// <summary>
        /// Generate PDF from parsed document using QuestPDF
        /// </summary>
        private byte[] GeneratePdfFromParsed(ParsedWordDocument doc)
        {
            return QuestPDF.Fluent.Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(50); // ~1.75 cm margin

                    // Add header if present
                    if (doc.HasHeader)
                    {
                        page.Header().Element(header =>
                        {
                            header.Column(column =>
                            {
                                foreach (var para in doc.Header!.Paragraphs)
                                {
                                    AddParagraphToColumn(column, para, true);
                                }
                            });
                        });
                    }

                    // Add content
                    page.Content().Element(content =>
                    {
                        content.Column(column =>
                        {
                            foreach (var element in doc.BodyElements)
                            {
                                if (element is ParagraphDocumentElement paraElement)
                                {
                                    AddParagraphToColumn(column, paraElement.Paragraph);
                                }
                                else if (element is TableDocumentElement tableElement)
                                {
                                    AddTableToColumn(column, tableElement.Table);
                                }
                            }
                        });
                    });

                    // Add footer if present
                    if (doc.HasFooter)
                    {
                        page.Footer().Element(footer =>
                        {
                            footer.Column(column =>
                            {
                                foreach (var para in doc.Footer!.Paragraphs)
                                {
                                    AddParagraphToColumn(column, para, true);
                                }

                                if (doc.Footer.HasPageNumber)
                                {
                                    column.Item().AlignCenter().Text(x =>
                                    {
                                        x.Span("Page ");
                                        x.CurrentPageNumber();
                                    });
                                }
                            });
                        });
                    }
                    else
                    {
                        // Default footer with page number
                        page.Footer().AlignCenter().Text(x =>
                        {
                            x.Span("Page ");
                            x.CurrentPageNumber();
                        });
                    }
                });
            }).GeneratePdf();
        }

        /// <summary>
        /// Add a paragraph to QuestPDF column
        /// </summary>
        private void AddParagraphToColumn(ColumnDescriptor column, ParagraphElement paragraph, bool isHeaderFooter = false)
        {
            IContainer textBuilder = column.Item();

            // Apply alignment
            textBuilder = paragraph.Alignment switch
            {
                Models.TextAlignment.Center => textBuilder.AlignCenter(),
                Models.TextAlignment.Right => textBuilder.AlignRight(),
                Models.TextAlignment.Justify => textBuilder.AlignLeft(),
                _ => textBuilder.AlignLeft()
            };

            // Apply spacing
            if (paragraph.Spacing != null)
            {
                var spacingTop = paragraph.Spacing.Before / 28.35;
                var spacingBottom = paragraph.Spacing.After / 28.35;
                textBuilder = textBuilder.PaddingTop((float)spacingTop).PaddingBottom((float)spacingBottom);
            }

            // Apply indentation
            if (paragraph.Indent != null)
            {
                var indentLeft = paragraph.Indent.Left / 28.35;
                var indentRight = paragraph.Indent.Right / 28.35;
                textBuilder = textBuilder.PaddingLeft((float)indentLeft).PaddingRight((float)indentRight);
            }

            // Skip empty paragraphs
            if (paragraph.Runs.Count == 0)
                return;

            var hasText = paragraph.Runs.Any(r => !string.IsNullOrWhiteSpace(r.Text));
            if (!hasText)
                return;

            // Render text with formatting
            textBuilder.Text(x =>
            {
                foreach (var run in paragraph.Runs)
                {
                    if (string.IsNullOrEmpty(run.Text))
                        continue;

                    var span = x.Span(run.Text);

                    if (run.Bold && !isHeaderFooter)
                        span.Bold();
                    if (run.Italic && !isHeaderFooter)
                        span.Italic();

                    if (run.FontSize != 11)
                        span.FontSize(run.FontSize);

                    if (!string.IsNullOrEmpty(run.Color))
                    {
                        var color = ParseColor(run.Color);
                        span.FontColor(color);
                    }
                }
            });
        }

        /// <summary>
        /// Check if all runs have the same formatting
        /// </summary>
        private bool AllRunsHaveSameFormatting(List<TextRun> runs)
        {
            if (runs.Count <= 1)
                return true;

            var first = runs[0];
            for (int i = 1; i < runs.Count; i++)
            {
                if (runs[i].Bold != first.Bold ||
                    runs[i].Italic != first.Italic ||
                    runs[i].FontSize != first.FontSize ||
                    runs[i].Color != first.Color)
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Add a table to QuestPDF column
        /// </summary>
        private void AddTableToColumn(ColumnDescriptor column, TableElement table)
        {
            column.Item().Table(tableDescriptor =>
            {
                // Define columns
                tableDescriptor.ColumnsDefinition(columns =>
                {
                    for (int i = 0; i < table.ColumnCount; i++)
                    {
                        columns.RelativeColumn();
                    }
                });

                // Add rows by iterating through cells
                foreach (var row in table.Rows)
                {
                    for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                    {
                        var cell = row.Cells[colIndex];

                        if (row.IsHeader)
                        {
                            tableDescriptor.Cell().Element(cellDescriptor =>
                            {
                                ApplyTableCellContent(cellDescriptor, cell, table.HasBorders);
                            });
                        }
                        else
                        {
                            tableDescriptor.Cell().Element(cellDescriptor =>
                            {
                                ApplyTableCellContent(cellDescriptor, cell, table.HasBorders);
                            });
                        }
                    }
                }
            });
        }

        private void ApplyTableCellContent(IContainer cellContainer, TableCellElement cell, bool hasBorders)
        {
            // Apply border if needed
            if (hasBorders)
            {
                cellContainer = cellContainer.Border(1);
            }

            // Apply alignment using the container methods
            cellContainer = cell.TextAlignment switch
            {
                Models.TextAlignment.Center => cellContainer.AlignCenter(),
                Models.TextAlignment.Right => cellContainer.AlignRight(),
                Models.TextAlignment.Justify => cellContainer.AlignLeft(),
                _ => cellContainer.AlignLeft()
            };

            // Apply vertical alignment
            cellContainer = cell.VerticalAlignment switch
            {
                Models.VerticalAlignment.Center => cellContainer.AlignMiddle(),
                Models.VerticalAlignment.Bottom => cellContainer.AlignBottom(),
                _ => cellContainer.AlignTop()
            };

            // Add content using padding
            cellContainer.Padding(5).Column(cellColumn =>
            {
                foreach (var para in cell.Paragraphs)
                {
                    AddParagraphToColumn(cellColumn, para);
                }
            });
        }

        /// <summary>
        /// Parse hex color string to QuestPDF color
        /// </summary>
        private string ParseColor(string hexColor)
        {
            if (string.IsNullOrEmpty(hexColor))
                return Colors.Black;

            // Remove # if present
            hexColor = hexColor.TrimStart('#');

            // Try to parse as hex
            if (hexColor.Length == 6)
            {
                return "#" + hexColor;
            }

            return Colors.Black;
        }
    }
}
