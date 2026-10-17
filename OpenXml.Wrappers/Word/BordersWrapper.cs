using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The borders of a table, a table row or a table cell, as <c>Borders</c> in Word.
/// </summary>
/// <remarks>
/// The indexer returns a <see cref="BorderWrapper"/> for a position. <see cref="BorderPosition.Horizontal"/> and
/// <see cref="BorderPosition.Vertical"/> are the lines between the rows and between the columns. The diagonals only exist for cells. A row has
/// no inside horizontal border (it is a single row) and its borders are written to its cells: see <see cref="BorderWrapper"/>.
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
            if ((_isTable || _owner is TableRow) && (position == BorderPosition.DiagonalDown || position == BorderPosition.DiagonalUp))
            {
                throw new NotSupportedException("Only a cell has diagonal borders.");
            }
            if (_owner is TableRow row)
            {
                if (position == BorderPosition.Horizontal)
                {
                    throw new NotSupportedException("A row has no horizontal inside border.");
                }
                return new BorderWrapper(GetRowTargets(row, position));
            }
            return new BorderWrapper(this, position);
        }
    }

    //The borders of a row are those of its cells: top and bottom on all of them, left on the first, right on the last, and the line between
    //the cells as the right border of one and the left border of the next.
    private static List<(BordersWrapper Borders, BorderPosition Position)> GetRowTargets(TableRow row, BorderPosition position)
    {
        List<TableCell> cells = row.Elements<TableCell>().ToList();
        List<(BordersWrapper, BorderPosition)> targets = new List<(BordersWrapper, BorderPosition)>();
        for (int i = 0; i < cells.Count; i++)
        {
            BordersWrapper borders = new BordersWrapper(cells[i], false);
            if (position == BorderPosition.Top || position == BorderPosition.Bottom
                || (position == BorderPosition.Left && i == 0) || (position == BorderPosition.Right && i == cells.Count - 1))
            {
                targets.Add((borders, position));
            }
            else if (position == BorderPosition.Vertical)
            {
                if (i > 0)
                {
                    targets.Add((borders, BorderPosition.Left));
                }
                if (i < cells.Count - 1)
                {
                    targets.Add((borders, BorderPosition.Right));
                }
            }
        }
        return targets;
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
        Set(_owner is TableRow ? new[] { BorderPosition.Vertical } : new[] { BorderPosition.Horizontal, BorderPosition.Vertical }, lineStyle, lineWidth, color);
    }

    /// <summary>
    /// Removes all the direct border settings.
    /// </summary>
    public void Clear()
    {
        if (_owner is TableRow row)
        {
            foreach (TableCell cell in row.Elements<TableCell>())
            {
                new BordersWrapper(cell, false).Clear();
            }
            return;
        }

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
