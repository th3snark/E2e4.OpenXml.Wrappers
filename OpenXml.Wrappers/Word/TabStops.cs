using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// How text lines up at a tab stop, as <c>WdTabAlignment</c> in Word (with the same values).
/// </summary>
public enum TabAlignment
{
    Left = 0,
    Center = 1,
    Right = 2,
    DecimalPoint = 3,
    Bar = 4,
    List = 6
}

/// <summary>
/// The character that fills the space before a tab stop, as <c>WdTabLeader</c> in Word (with the same values).
/// </summary>
public enum TabLeader
{
    Spaces = 0,
    Dots = 1,
    Dashes = 2,
    Lines = 3,
    Heavy = 4,
    MiddleDot = 5
}

/// <summary>
/// A tab stop of a paragraph, as <c>TabStop</c> in Word.
/// </summary>
/// <remarks>
/// The tab stop is found again by its position each time a property is used, so it belongs to the paragraphs of the format that it came from
/// and stays valid while the position is not changed by something else. Properties read the tab stops that apply, including those that come
/// from styles.
/// </remarks>
public class TabStopWrapper
{
    private readonly ParagraphFormatWrapper _format;
    private long _twips;

    internal TabStopWrapper(ParagraphFormatWrapper format, long twips)
    {
        _format = format;
        _twips = twips;
    }

    private TabStop Find(ParagraphProperties properties)
    {
        return properties?.Tabs?.Elements<TabStop>().FirstOrDefault(t => t.Position != null && t.Position.HasValue && t.Position.Value == _twips && !IsClear(t));
    }

    internal static bool IsClear(TabStop tab)
    {
        return tab.Val != null && tab.Val.HasValue && tab.Val.Value == TabStopValues.Clear;
    }

    /// <summary>The position of the tab stop in points from the left margin. Setting it moves the tab stop.</summary>
    public double Position
    {
        get { return _twips / 20.0; }
        set
        {
            long twips = TabStopWrapperCollection.ToTwips(value);
            long previous = _twips;
            _format.Apply(properties =>
            {
                TabStop tab = properties.Tabs?.Elements<TabStop>().FirstOrDefault(t => t.Position != null && t.Position.HasValue && t.Position.Value == previous && !IsClear(t));
                if (tab != null)
                {
                    tab.Position = (int)twips;
                    TabStopWrapperCollection.Sort(properties.Tabs);
                }
            });
            _twips = twips;
        }
    }

    /// <summary>How text lines up at the tab stop.</summary>
    public TabAlignment Alignment
    {
        get
        {
            return _format.GetValue(properties =>
            {
                TabStop tab = Find(properties);
                TabStopValues? value = tab?.Val?.Value;
                if (value == TabStopValues.Center)
                {
                    return TabAlignment.Center;
                }
                if (value == TabStopValues.Right || value == TabStopValues.End)
                {
                    return TabAlignment.Right;
                }
                if (value == TabStopValues.Decimal)
                {
                    return TabAlignment.DecimalPoint;
                }
                if (value == TabStopValues.Bar)
                {
                    return TabAlignment.Bar;
                }
                return value == TabStopValues.Number ? TabAlignment.List : TabAlignment.Left;
            });
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            _format.Apply(properties =>
            {
                TabStop tab = Find(properties);
                if (tab != null)
                {
                    tab.Val = TabStopWrapperCollection.ToTabStopValues(value);
                }
            });
        }
    }

    /// <summary>The character that fills the space before the tab stop.</summary>
    public TabLeader Leader
    {
        get
        {
            return _format.GetValue(properties =>
            {
                TabStopLeaderCharValues? value = Find(properties)?.Leader?.Value;
                if (value == TabStopLeaderCharValues.Dot)
                {
                    return TabLeader.Dots;
                }
                if (value == TabStopLeaderCharValues.Hyphen)
                {
                    return TabLeader.Dashes;
                }
                if (value == TabStopLeaderCharValues.Underscore)
                {
                    return TabLeader.Lines;
                }
                if (value == TabStopLeaderCharValues.Heavy)
                {
                    return TabLeader.Heavy;
                }
                return value == TabStopLeaderCharValues.MiddleDot ? TabLeader.MiddleDot : TabLeader.Spaces;
            });
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            _format.Apply(properties =>
            {
                TabStop tab = Find(properties);
                if (tab != null)
                {
                    tab.Leader = TabStopWrapperCollection.ToLeader(value);
                }
            });
        }
    }

    /// <summary>
    /// Removes the tab stop. A tab stop that comes from a style is switched off for the paragraphs instead, as Word does.
    /// </summary>
    public void Clear()
    {
        long twips = _twips;
        _format.Apply(properties =>
        {
            Tabs tabs = properties.Tabs;
            TabStop direct = tabs?.Elements<TabStop>().FirstOrDefault(t => t.Position != null && t.Position.HasValue && t.Position.Value == twips && !IsClear(t));
            if (direct != null)
            {
                direct.Remove();
            }
            if (_format.IsInherited(twips))
            {
                if (properties.Tabs == null)
                {
                    properties.Tabs = new Tabs();
                }
                properties.Tabs.AppendChild(new TabStop() { Val = TabStopValues.Clear, Position = (int)twips });
                TabStopWrapperCollection.Sort(properties.Tabs);
            }
            if (properties.Tabs != null && !properties.Tabs.HasChildren)
            {
                properties.Tabs.Remove();
            }
        });
    }
}

/// <summary>
/// The tab stops of a paragraph format, as <c>TabStops</c> in Word.
/// </summary>
/// <remarks>
/// The collection lists the tab stops that apply to the paragraphs: those set directly and those that come from the styles. It is empty when
/// the paragraphs of the format do not have the same tab stops. Adding a tab stop adds it to every paragraph.
/// </remarks>
public class TabStopWrapperCollection : IEnumerable<TabStopWrapper>
{
    private readonly ParagraphFormatWrapper _format;

    internal TabStopWrapperCollection(ParagraphFormatWrapper format)
    {
        _format = format;
    }

    //The positions of the tab stops that apply, in order; the same for every paragraph or none.
    private IList<long> GetPositions()
    {
        IList<long> positions = _format.GetValue<IList<long>>(properties => properties?.Tabs == null
            ? new List<long>()
            : properties.Tabs.Elements<TabStop>().Where(t => !TabStopWrapper.IsClear(t) && t.Position != null && t.Position.HasValue).Select(t => (long)t.Position.Value).OrderBy(p => p).ToList(),
            new ListComparer());
        return positions ?? new List<long>();
    }

    private sealed class ListComparer : IEqualityComparer<IList<long>>
    {
        public bool Equals(IList<long> x, IList<long> y)
        {
            return x.SequenceEqual(y);
        }

        public int GetHashCode(IList<long> obj)
        {
            return obj.Count;
        }
    }

    /// <summary>The number of tab stops.</summary>
    public int Count
    {
        get { return GetPositions().Count; }
    }

    /// <summary>The tab stop at the index, in order of position.</summary>
    public TabStopWrapper this[int index]
    {
        get
        {
            IList<long> positions = GetPositions();
            if (index < 0 || index >= positions.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return new TabStopWrapper(_format, positions[index]);
        }
    }

    /// <summary>
    /// Adds a tab stop at the position (in points from the left margin) to every paragraph, or changes the one that is already there.
    /// </summary>
    public TabStopWrapper Add(double position, TabAlignment alignment = TabAlignment.Left, TabLeader leader = TabLeader.Spaces)
    {
        if (!EnumCompat.IsDefined(alignment))
        {
            throw new ArgumentOutOfRangeException(nameof(alignment));
        }
        if (!EnumCompat.IsDefined(leader))
        {
            throw new ArgumentOutOfRangeException(nameof(leader));
        }

        long twips = ToTwips(position);
        _format.Apply(properties =>
        {
            if (properties.Tabs == null)
            {
                properties.Tabs = new Tabs();
            }
            foreach (TabStop existing in properties.Tabs.Elements<TabStop>().Where(t => t.Position != null && t.Position.HasValue && t.Position.Value == twips).ToList())
            {
                existing.Remove();
            }
            TabStop tab = new TabStop() { Val = ToTabStopValues(alignment), Position = (int)twips };
            if (leader != TabLeader.Spaces)
            {
                tab.Leader = ToLeader(leader);
            }
            properties.Tabs.AppendChild(tab);
            Sort(properties.Tabs);
        });
        return new TabStopWrapper(_format, twips);
    }

    /// <summary>Removes every tab stop, including those that come from styles (which are switched off for the paragraphs).</summary>
    public void ClearAll()
    {
        foreach (long position in GetPositions().ToList())
        {
            new TabStopWrapper(_format, position).Clear();
        }
    }

    internal static long ToTwips(double points)
    {
        if (double.IsNaN(points) || points < 0 || points > 1584)
        {
            throw new ArgumentOutOfRangeException(nameof(points), "The position must be between 0 and 1584 points.");
        }
        return (long)Math.Round(points * 20, MidpointRounding.AwayFromZero);
    }

    internal static void Sort(Tabs tabs)
    {
        if (tabs == null)
        {
            return;
        }
        List<TabStop> sorted = tabs.Elements<TabStop>().OrderBy(t => t.Position != null && t.Position.HasValue ? t.Position.Value : 0).ToList();
        foreach (TabStop tab in sorted)
        {
            tab.Remove();
        }
        foreach (TabStop tab in sorted)
        {
            tabs.AppendChild(tab);
        }
    }

    internal static TabStopValues ToTabStopValues(TabAlignment alignment)
    {
        return alignment == TabAlignment.Center ? TabStopValues.Center
            : alignment == TabAlignment.Right ? TabStopValues.Right
            : alignment == TabAlignment.DecimalPoint ? TabStopValues.Decimal
            : alignment == TabAlignment.Bar ? TabStopValues.Bar
            : alignment == TabAlignment.List ? TabStopValues.Number
            : TabStopValues.Left;
    }

    internal static TabStopLeaderCharValues ToLeader(TabLeader leader)
    {
        return leader == TabLeader.Dots ? TabStopLeaderCharValues.Dot
            : leader == TabLeader.Dashes ? TabStopLeaderCharValues.Hyphen
            : leader == TabLeader.Lines ? TabStopLeaderCharValues.Underscore
            : leader == TabLeader.Heavy ? TabStopLeaderCharValues.Heavy
            : leader == TabLeader.MiddleDot ? TabStopLeaderCharValues.MiddleDot
            : TabStopLeaderCharValues.None;
    }

    #region IEnumerable Members

    public IEnumerator<TabStopWrapper> GetEnumerator()
    {
        foreach (long position in GetPositions())
        {
            yield return new TabStopWrapper(_format, position);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
