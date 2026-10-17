using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The borders of a table or a table cell, as <c>Borders</c> in Word.
/// </summary>
/// <remarks>
/// The indexer returns a <see cref="BorderWrapper"/> for a position. <see cref="BorderPosition.Horizontal"/> and
/// <see cref="BorderPosition.Vertical"/> are the lines between the rows and between the columns. The diagonals only exist for cells.
/// </remarks>
public class BordersWrapper
{
    //The order of the border elements in the schema, which is the same in table and cell borders.
    private static readonly string[] SchemaOrder = { "top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl" };

    private readonly OpenXmlElement _owner;
    private readonly bool _isTable;

    internal BordersWrapper(OpenXmlElement owner, bool isTable)
    {
        _owner = owner;
        _isTable = isTable;
    }

    public BorderWrapper this[BorderPosition position]
    {
        get
        {
            if (!Enum.IsDefined(position))
            {
                throw new ArgumentOutOfRangeException(nameof(position));
            }
            if (_isTable && (position == BorderPosition.DiagonalDown || position == BorderPosition.DiagonalUp))
            {
                throw new NotSupportedException("A table does not have diagonal borders.");
            }
            return new BorderWrapper(this, position);
        }
    }

    /// <summary>
    /// Sets the style, width and colour of the four outside borders.
    /// </summary>
    public void SetOutside(BorderLineStyle lineStyle, double lineWidth, string color)
    {
        Set(new[] { BorderPosition.Top, BorderPosition.Left, BorderPosition.Bottom, BorderPosition.Right }, lineStyle, lineWidth, color);
    }

    /// <summary>
    /// Sets the style, width and colour of the borders between the rows and between the columns.
    /// </summary>
    public void SetInside(BorderLineStyle lineStyle, double lineWidth, string color)
    {
        Set(new[] { BorderPosition.Horizontal, BorderPosition.Vertical }, lineStyle, lineWidth, color);
    }

    /// <summary>
    /// Removes all the direct border settings.
    /// </summary>
    public void Clear()
    {
        OpenXmlCompositeElement container = GetContainer(false);
        container?.Remove();
    }

    private void Set(BorderPosition[] positions, BorderLineStyle lineStyle, double lineWidth, string color)
    {
        foreach (BorderPosition position in positions)
        {
            BorderWrapper border = this[position];
            border.LineStyle = lineStyle;
            if (lineStyle != BorderLineStyle.None)
            {
                border.LineWidth = lineWidth;
                border.Color = color;
            }
        }
    }

    private OpenXmlCompositeElement GetContainer(bool create)
    {
        if (_isTable)
        {
            Table table = (Table)_owner;
            TableProperties properties = table.GetFirstChild<TableProperties>();
            if (properties == null && create)
            {
                properties = new TableProperties();
                table.PrependChild(properties);
            }
            if (properties?.TableBorders == null && create)
            {
                properties.TableBorders = new TableBorders();
            }
            return properties?.TableBorders;
        }

        TableCell cell = (TableCell)_owner;
        TableCellProperties cellProperties = cell.TableCellProperties;
        if (cellProperties == null && create)
        {
            cellProperties = new TableCellProperties();
            cell.PrependChild(cellProperties);
        }
        if (cellProperties?.TableCellBorders == null && create)
        {
            cellProperties.TableCellBorders = new TableCellBorders();
        }
        return cellProperties?.TableCellBorders;
    }

    private static string[] GetNames(BorderPosition position)
    {
        return position switch
        {
            BorderPosition.Top => new[] { "top" },
            BorderPosition.Left => new[] { "left", "start" },
            BorderPosition.Bottom => new[] { "bottom" },
            BorderPosition.Right => new[] { "right", "end" },
            BorderPosition.Horizontal => new[] { "insideH" },
            BorderPosition.Vertical => new[] { "insideV" },
            BorderPosition.DiagonalDown => new[] { "tl2br" },
            _ => new[] { "tr2bl" }
        };
    }

    internal BorderType Find(BorderPosition position)
    {
        string[] names = GetNames(position);
        return GetContainer(false)?.Elements<BorderType>().FirstOrDefault(e => names.Contains(e.LocalName));
    }

    internal BorderType GetOrAdd(BorderPosition position)
    {
        BorderType border = Find(position);
        if (border != null)
        {
            return border;
        }

        OpenXmlCompositeElement container = GetContainer(true);
        border = position switch
        {
            BorderPosition.Top => new TopBorder(),
            BorderPosition.Left => new LeftBorder(),
            BorderPosition.Bottom => new BottomBorder(),
            BorderPosition.Right => new RightBorder(),
            BorderPosition.Horizontal => new InsideHorizontalBorder(),
            BorderPosition.Vertical => new InsideVerticalBorder(),
            BorderPosition.DiagonalDown => new TopLeftToBottomRightCellBorder(),
            _ => new TopRightToBottomLeftCellBorder()
        };

        //Keep the schema order: go before the first existing border that comes later.
        int rank = Array.IndexOf(SchemaOrder, border.LocalName);
        OpenXmlElement next = container.ChildElements.FirstOrDefault(e => Array.IndexOf(SchemaOrder, e.LocalName) > rank);
        if (next != null)
        {
            next.InsertBeforeSelf(border);
        }
        else
        {
            container.AppendChild(border);
        }
        return border;
    }
}
