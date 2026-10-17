using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// Adds a new empty paragraph after the range and returns it, as <c>Range.InsertParagraphAfter</c> in Word (which returns nothing).
    /// </summary>
    /// <remarks>
    /// The paragraph is a plain one: it does not take on the formatting of the paragraph before it. After a paragraph, table or bookmark or
    /// comment range it follows the paragraph or table; at the end of a table cell it is the last paragraph of the cell; and after the
    /// document it is the last paragraph of the body. This is the way to append to the end of a document that ends in a table, where
    /// <see cref="ParagraphWrapperCollection.Add"/> would add the paragraph to the last cell of the table.
    /// </remarks>
    /// <exception cref="NotSupportedException">The range is of a type that does not support adding paragraphs, such as a table row.</exception>
    public ParagraphWrapper InsertParagraphAfter()
    {
        return InsertParagraph(true);
    }

    /// <summary>
    /// Adds a new empty paragraph before the range and returns it, as <c>Range.InsertParagraphBefore</c> in Word. See
    /// <see cref="InsertParagraphAfter"/>.
    /// </summary>
    /// <exception cref="NotSupportedException">The range is of a type that does not support adding paragraphs, such as a table row.</exception>
    public ParagraphWrapper InsertParagraphBefore()
    {
        return InsertParagraph(false);
    }

    private ParagraphWrapper InsertParagraph(bool after)
    {
        Paragraph paragraph = new Paragraph();
        if (_startOpenXmlElement is Paragraph || _startOpenXmlElement is Table)
        {
            AddBlockNextTo(_startOpenXmlElement, paragraph, after);
        }
        else if (_startOpenXmlElement is BookmarkStart || _startOpenXmlElement is CommentRangeStart)
        {
            //The paragraph of the marker that is nearer the new paragraph.
            OpenXmlElement marker = after ? _endOpenXmlElement : _startOpenXmlElement;
            OpenXmlElement block = marker.Ancestors<Paragraph>().FirstOrDefault();
            if (block == null)
            {
                throw new NotSupportedException("Adding a paragraph is only supported when the markers of the range are inside a paragraph.");
            }
            AddBlockNextTo(block, paragraph, after);
        }
        else if (_startOpenXmlElement is TableCell || _startOpenXmlElement is Document || _startOpenXmlElement is Body
            || _startOpenXmlElement is Header || _startOpenXmlElement is Footer)
        {
            OpenXmlCompositeElement container = _startOpenXmlElement is Document ? GetBody() : (OpenXmlCompositeElement)_startOpenXmlElement;
            AddToContainer(container, paragraph, after);
        }
        else
        {
            throw new NotSupportedException(String.Format(CultureInfo.InvariantCulture, "Adding a paragraph is not currently supported for a range starting with the type {0}", _startOpenXmlElement.GetType().Name));
        }
        return new ParagraphWrapper(paragraph);
    }

    private static void AddBlockNextTo(OpenXmlElement block, Paragraph paragraph, bool after)
    {
        if (after)
        {
            block.InsertAfterSelf(paragraph);
        }
        else
        {
            block.InsertBeforeSelf(paragraph);
        }
    }

    private static void AddToContainer(OpenXmlCompositeElement container, Paragraph paragraph, bool after)
    {
        //The section properties are not content, so new content goes before them.
        OpenXmlElement sectionProperties = container.Elements<SectionProperties>().LastOrDefault();
        if (!after)
        {
            //A table cell starts with its properties.
            OpenXmlElement cellProperties = container.Elements<TableCellProperties>().FirstOrDefault();
            if (cellProperties != null)
            {
                cellProperties.InsertAfterSelf(paragraph);
            }
            else
            {
                container.InsertAt(paragraph, 0);
            }
        }
        else if (sectionProperties != null)
        {
            sectionProperties.InsertBeforeSelf(paragraph);
        }
        else
        {
            container.AppendChild(paragraph);
        }
    }
}
