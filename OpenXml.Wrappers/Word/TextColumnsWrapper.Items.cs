using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The direction text flows from one column to the next, as <c>WdFlowDirection</c> in Word (with the same values).
/// </summary>
public enum FlowDirection
{
    /// <summary>From left to right.</summary>
    Ltr = 0,

    /// <summary>From right to left.</summary>
    Rtl = 1
}

public partial class TextColumnsWrapper : IEnumerable<TextColumnWrapper>
{
    private const int MaxColumns = 45;

    internal static string Twips(double points)
    {
        return ((uint)Math.Round(points * 20, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);
    }

    //Word keeps whole twips, dropping the remainder, when it divides the width of the text between columns.
    internal static uint EvenTwips(double textWidth, double spacing, int count)
    {
        return (uint)Math.Floor((textWidth - spacing * (count - 1)) / count * 20 + 0.001);
    }

    /// <summary>
    /// Lists every column one by one, which is how columns that differ in width are described, giving each the width it has now. Word keeps the
    /// number of columns when it does this.
    /// </summary>
    internal Columns Materialize()
    {
        Columns columns = _setup.GetOrAddProperty<Columns>();
        if (columns.Elements<Column>().Any())
        {
            return columns;
        }

        int count = Count;
        double spacing = Spacing;
        string each = EvenTwips(_setup.GetTextWidth(), spacing, count).ToString(CultureInfo.InvariantCulture);
        //Word writes 0 where the SDK would write false.
        columns.SetAttribute(new OpenXmlAttribute("w", "equalWidth", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0"));
        columns.ColumnCount = (short)count;
        for (int i = 0; i < count; i++)
        {
            Column column = new Column() { Width = each };
            if (i < count - 1)
            {
                column.Space = Twips(spacing);
            }
            columns.AppendChild(column);
        }
        return columns;
    }

    internal Column GetColumnElement(int index)
    {
        if (index < 0 || index >= Count)
        {
            throw new InvalidOperationException("The column does not exist.");
        }
        return Existing?.Elements<Column>().ElementAtOrDefault(index);
    }

    /// <summary>The column at the index (0-based), as <c>TextColumns.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no column at the index.</exception>
    public TextColumnWrapper this[int index]
    {
        get
        {
            if (index < 0 || index >= Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new TextColumnWrapper(this, index);
        }
    }

    /// <summary>
    /// Adds a column at the right, as <c>TextColumns.Add</c> in Word: it takes <paramref name="width"/> points, and the column before it
    /// is given <paramref name="spacing"/> points after it (36 by default). The first column takes whatever width of the text is left. Without a
    /// width the columns stay evenly spaced, and there is one more of them.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The width or the spacing is outside what a page can hold.</exception>
    /// <exception cref="ArgumentException">There is no width and the columns are not evenly spaced.</exception>
    /// <exception cref="InvalidOperationException">There are 45 columns already.</exception>
    public TextColumnWrapper Add(double? width = null, double? spacing = null)
    {
        if (width.HasValue && !(width.Value > 0 && width.Value <= 1584))
        {
            throw new ArgumentOutOfRangeException(nameof(width), "The width must be more than 0 and at most 1584 points.");
        }
        if (spacing.HasValue && !(spacing.Value >= 0 && spacing.Value <= 1584))
        {
            throw new ArgumentOutOfRangeException(nameof(spacing), "The spacing must be between 0 and 1584 points.");
        }

        int count = Count;
        if (count >= MaxColumns)
        {
            throw new InvalidOperationException("The text has the most columns, 45, already.");
        }
        if (!width.HasValue)
        {
            if (!EvenlySpaced)
            {
                throw new ArgumentException("The width of the column is needed when the columns are not all the same width.", nameof(width));
            }
            SetCount(count + 1);
            if (spacing.HasValue)
            {
                Spacing = spacing.Value;
            }
            return this[count];
        }

        Columns columns = Materialize();
        List<Column> list = columns.Elements<Column>().ToList();
        Column before = list[list.Count - 1];
        //Without spacing the new column is as far from the one before as that one is from the one before it, or 36 points if there is none.
        string previous = list.Count > 1 ? list[list.Count - 2].Space?.Value : null;
        before.Space = spacing.HasValue ? Twips(spacing.Value) : previous ?? "720";
        Column added = new Column() { Width = Twips(width.Value) };
        columns.AppendChild(added);
        list.Add(added);
        columns.ColumnCount = (short)list.Count;
        GiveFirstTheRest(list);
        return this[count];
    }

    //The first column is the one that takes up what the others leave of the width of the text.
    internal void GiveFirstTheRest(List<Column> list)
    {
        long rest = (long)Math.Round(_setup.GetTextWidth() * 20);
        for (int i = 1; i < list.Count; i++)
        {
            rest -= long.Parse(list[i].Width.Value, CultureInfo.InvariantCulture);
        }
        for (int i = 0; i < list.Count - 1; i++)
        {
            rest -= list[i].Space == null ? 0 : long.Parse(list[i].Space.Value, CultureInfo.InvariantCulture);
        }
        list[0].Width = Math.Max(rest, 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// The direction text flows from one column to the next, as <c>TextColumns.FlowDirection</c> in Word. It is the direction of the section, as
    /// <see cref="PageSetupWrapper.SectionDirection"/> has it.
    /// </summary>
    public FlowDirection FlowDirection
    {
        get { return _setup.SectionDirection == SectionDirection.Rtl ? FlowDirection.Rtl : FlowDirection.Ltr; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            _setup.SectionDirection = value == FlowDirection.Rtl ? SectionDirection.Rtl : SectionDirection.Ltr;
        }
    }

    /// <inheritdoc/>
    public IEnumerator<TextColumnWrapper> GetEnumerator()
    {
        int count = Count;
        for (int i = 0; i < count; i++)
        {
            yield return new TextColumnWrapper(this, i);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

/// <summary>
/// One column of text of a section, as <c>TextColumn</c> in Word. Changing the width or the spacing of a column lists the columns one by one, so
/// they are no longer evenly spaced.
/// </summary>
public class TextColumnWrapper
{
    private readonly TextColumnsWrapper _columns;
    private readonly int _index;

    internal TextColumnWrapper(TextColumnsWrapper columns, int index)
    {
        _columns = columns;
        _index = index;
    }

    /// <summary>The width of the column in points.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is not more than 0 and at most 1584 points.</exception>
    public double Width
    {
        get
        {
            Column column = _columns.GetColumnElement(_index);
            string width = column?.Width?.Value;
            if (width != null && double.TryParse(width, NumberStyles.Float, CultureInfo.InvariantCulture, out double twips))
            {
                return twips / 20;
            }
            return _columns.Width;
        }
        set
        {
            if (!(value > 0 && value <= 1584))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The width must be more than 0 and at most 1584 points.");
            }
            _columns.GetColumnElement(_index);
            _columns.Materialize().Elements<Column>().ElementAt(_index).Width = TextColumnsWrapper.Twips(value);
        }
    }

    /// <summary>The space after the column in points; 0 for the last column, which cannot be given any.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The spacing is not between 0 and 1584 points.</exception>
    /// <exception cref="InvalidOperationException">The column is the last.</exception>
    public double SpaceAfter
    {
        get
        {
            if (_index == _columns.Count - 1)
            {
                return 0;
            }
            string space = _columns.GetColumnElement(_index)?.Space?.Value;
            return space != null && double.TryParse(space, NumberStyles.Float, CultureInfo.InvariantCulture, out double twips) ? twips / 20 : _columns.Spacing;
        }
        set
        {
            if (double.IsNaN(value) || value < 0 || value > 1584)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The spacing must be between 0 and 1584 points.");
            }
            _columns.GetColumnElement(_index);
            if (_index == _columns.Count - 1)
            {
                throw new InvalidOperationException("The last column has no space after it.");
            }

            _columns.Materialize().Elements<Column>().ElementAt(_index).Space = TextColumnsWrapper.Twips(value);
        }
    }
}
