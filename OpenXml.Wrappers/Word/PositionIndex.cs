using System;
using System.Linq;

using DocumentFormat.OpenXml;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The positions of the text of one story, worked out once so that the positions of many ranges can be read quickly.
/// </summary>
/// <remarks>
/// <para>Reading <see cref="RangeWrapper.Start"/> or <see cref="RangeWrapper.End"/> counts through the story up to the range, which is quick
/// for one range but slow when done for thousands in a big document. An index is built by one pass over the story (<see cref="Create"/>) and
/// then gives the position of any range in it at once.</para>
/// <para>An index is a snapshot: it describes the story as it was when it was built. Build a new one (or call <see cref="Refresh"/>) after
/// text has been added or removed, or the positions it gives will be wrong. The ranges themselves are not affected, they follow their text.</para>
/// </remarks>
public sealed class PositionIndex
{
    private readonly OpenXmlElement _story;
    private StoryMap _map;

    private PositionIndex(OpenXmlElement story)
    {
        _story = story;
        _map = StoryMap.Build(story);
    }

    /// <summary>
    /// Builds the index of the story that the range is in.
    /// </summary>
    public static PositionIndex Create(RangeWrapper range)
    {
        ArgumentNullException.ThrowIfNull(range);
        return new PositionIndex(StoryMap.FindRoot(range.AnchorElement));
    }

    /// <summary>
    /// The number of positions in the story.
    /// </summary>
    public int Length
    {
        get { return _map.Length; }
    }

    /// <summary>
    /// Builds the index again, for the story as it is now.
    /// </summary>
    public void Refresh()
    {
        _map = StoryMap.Build(_story);
    }

    /// <summary>
    /// The position of the start of the range.
    /// </summary>
    /// <exception cref="ArgumentException">The range is not in the story of the index.</exception>
    public int StartOf(RangeWrapper range)
    {
        return PositionsOf(range).Start;
    }

    /// <summary>
    /// The position of the end of the range.
    /// </summary>
    /// <exception cref="ArgumentException">The range is not in the story of the index.</exception>
    public int EndOf(RangeWrapper range)
    {
        return PositionsOf(range).End;
    }

    /// <summary>
    /// The positions of the start and end of the range.
    /// </summary>
    /// <exception cref="ArgumentException">The range is not in the story of the index.</exception>
    public (int Start, int End) PositionsOf(RangeWrapper range)
    {
        ArgumentNullException.ThrowIfNull(range);
        if (!ReferenceEquals(StoryMap.FindRoot(range.AnchorElement), _story))
        {
            throw new ArgumentException("The range is not in the story that the index is of.", nameof(range));
        }

        if (range.IsSpan)
        {
            int start = RangeWrapper.ResolveLocal(_map, range.StartAnchor);
            int end = RangeWrapper.ResolveLocal(_map, range.EndAnchor);
            if (range.EndAnchor.Element == null)
            {
                end = _map.Length;
            }
            if (start >= 0 && end >= 0)
            {
                return (start, Math.Max(start, end));
            }
        }
        else
        {
            //A range of elements is worked out in the part of the story it is in, which starts at the first piece of text there.
            range.GetScoped(false, out StoryMap scope, out int scopeStart, out int scopeEnd);
            if (scope.IsStory)
            {
                return (scopeStart, scopeEnd);
            }
            if (scope.Atoms.Count > 0 && _map.TryGetAtom(scope.Atoms[0].Element, out Atom first))
            {
                return (first.Start + scopeStart, first.Start + scopeEnd);
            }
        }

        //The text the range was on is not in this snapshot, so the range has to work its position out for itself.
        return (range.Start, range.End);
    }
}
