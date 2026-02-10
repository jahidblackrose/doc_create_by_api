using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BankingDocumentAPI.Models
{
    /// <summary>
    /// Represents the type of document element
    /// </summary>
    public enum ElementType
    {
        Paragraph,
        Table,
        Image
    }

    /// <summary>
    /// Represents text alignment options
    /// </summary>
    public enum TextAlignment
    {
        Left,
        Center,
        Right,
        Justify
    }

    /// <summary>
    /// Represents a text run with formatting
    /// </summary>
    public class TextRun
    {
        public string Text { get; set; } = string.Empty;
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string? FontFamily { get; set; }
        public int FontSize { get; set; } = 11; // in points
        public string? Color { get; set; } // Hex color

        public TextRun()
        {
        }

        public TextRun(string text)
        {
            Text = text;
        }
    }

    /// <summary>
    /// Represents spacing information for paragraphs
    /// </summary>
    public class SpacingInfo
    {
        public double Before { get; set; } // in points
        public double After { get; set; } // in points
        public double LineSpacing { get; set; } = 1.0; // multiplier
    }

    /// <summary>
    /// Represents indentation information for paragraphs
    /// </summary>
    public class IndentInfo
    {
        public double Left { get; set; } // in points
        public double Right { get; set; } // in points
        public double FirstLine { get; set; } // in points
    }

    /// <summary>
    /// Represents a paragraph element with runs and formatting
    /// </summary>
    public class ParagraphElement
    {
        public List<TextRun> Runs { get; set; } = new();
        public TextAlignment Alignment { get; set; } = TextAlignment.Left;
        public SpacingInfo? Spacing { get; set; }
        public IndentInfo? Indent { get; set; }
        public bool IsHeading { get; set; }
        public int? HeadingLevel { get; set; }
    }

    /// <summary>
    /// Represents a table cell
    /// </summary>
    public class TableCellElement
    {
        public List<ParagraphElement> Paragraphs { get; set; } = new();
        public int? ColumnSpan { get; set; }
        public int? RowSpan { get; set; }
        public string? BackgroundColor { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Top;
        public TextAlignment TextAlignment { get; set; } = TextAlignment.Left;
        public double? CellWidth { get; set; } // in points
    }

    /// <summary>
    /// Represents vertical alignment for table cells
    /// </summary>
    public enum VerticalAlignment
    {
        Top,
        Center,
        Bottom
    }

    /// <summary>
    /// Represents a table row
    /// </summary>
    public class TableRowElement
    {
        public List<TableCellElement> Cells { get; set; } = new();
        public double? Height { get; set; } // in points
        public bool IsHeader { get; set; }
    }

    /// <summary>
    /// Represents a table element
    /// </summary>
    public class TableElement
    {
        public List<TableRowElement> Rows { get; set; } = new();
        public int ColumnCount { get; set; }
        public bool HasBorders { get; set; } = true;
        public double? TableWidth { get; set; } // in points
    }

    /// <summary>
    /// Represents an image element
    /// </summary>
    public class ImageElement
    {
        public byte[] ImageData { get; set; } = Array.Empty<byte>();
        public string? ContentType { get; set; }
        public double Width { get; set; } // in points
        public double Height { get; set; } // in points
        public TextAlignment Alignment { get; set; } = TextAlignment.Left;
    }

    /// <summary>
    /// Represents a header or footer element
    /// </summary>
    public class HeaderFooterElement
    {
        public List<ParagraphElement> Paragraphs { get; set; } = new();
        public bool HasPageNumber { get; set; }
    }

    /// <summary>
    /// Represents a parsed document element (base class)
    /// </summary>
    public abstract class DocumentElement
    {
        public ElementType Type { get; set; }
    }

    /// <summary>
    /// Represents a paragraph document element
    /// </summary>
    public class ParagraphDocumentElement : DocumentElement
    {
        public ParagraphElement Paragraph { get; set; } = new();

        public ParagraphDocumentElement()
        {
            Type = ElementType.Paragraph;
        }
    }

    /// <summary>
    /// Represents a table document element
    /// </summary>
    public class TableDocumentElement : DocumentElement
    {
        public TableElement Table { get; set; } = new();

        public TableDocumentElement()
        {
            Type = ElementType.Table;
        }
    }

    /// <summary>
    /// Represents an image document element
    /// </summary>
    public class ImageDocumentElement : DocumentElement
    {
        public ImageElement Image { get; set; } = new();

        public ImageDocumentElement()
        {
            Type = ElementType.Image;
        }
    }

    /// <summary>
    /// Represents the complete parsed Word document
    /// </summary>
    public class ParsedWordDocument
    {
        public List<DocumentElement> BodyElements { get; set; } = new();
        public HeaderFooterElement? Header { get; set; }
        public HeaderFooterElement? Footer { get; set; }
        public bool HasHeader => Header != null && (Header.Paragraphs.Count > 0 || Header.HasPageNumber);
        public bool HasFooter => Footer != null && (Footer.Paragraphs.Count > 0 || Footer.HasPageNumber);
    }
}
