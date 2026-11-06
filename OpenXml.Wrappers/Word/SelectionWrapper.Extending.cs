using System;
using System.Linq;

namespace OpenXml.Wrappers.Word;

public partial class SelectionWrapper
{
    //Extend and Shrink grow and shrink the selection by units around a place: the one the selection was extended from. What the last of
    //them left is remembered, so that the place is known again as long as the selection is as they left it.
    private int _ladderAnchor;
    private int _ladderStart = -1;
    private int _ladderEnd = -1;

    private static readonly TextUnit[] Ladder = { TextUnit.Word, TextUnit.Sentence, TextUnit.Paragraph, TextUnit.Story };
    private static readonly TextUnit[] LadderDown = { TextUnit.Story, TextUnit.Paragraph, TextUnit.Sentence, TextUnit.Word };

    /// <summary>
    /// True while the selection is being extended, as <c>Selection.ExtendMode</c> in Word: the movement methods then extend the selection
    /// instead of moving the insertion point. Typing turns it off.
    /// </summary>
    public bool ExtendMode { get; set; }

    private int Anchor()
    {
        return _ladderStart == _range.Start && _ladderEnd == _range.End ? _ladderAnchor : _range.Start;
    }

    private void SetLadder(int start, int end, int anchor)
    {
        SetRange(start, end);
        _ladderAnchor = anchor;
        _ladderStart = start;
        _ladderEnd = end;
    }

    //The unit of the ladder (word, sentence, paragraph, story) around a place.
    private RangeWrapper UnitAround(int position, TextUnit unit)
    {
        RangeWrapper range = _range.Duplicate();
        range.SetRange(position, position);
        range.Expand(unit);
        return range;
    }

    //An extending movement is one while extend mode is on.
    private Movement Effective(Movement extend)
    {
        return ExtendMode && extend == Movement.Move ? Movement.Extend : extend;
    }

    /// <summary>
    /// Turns extend mode on, or when it is on already makes the selection the next larger unit around the place it was extended from
    /// (a word, a sentence, a paragraph, the story), as <c>Selection.Extend</c> in Word without an argument. The call that turns the
    /// mode on selects nothing.
    /// </summary>
    public void Extend()
    {
        if (!ExtendMode)
        {
            ExtendMode = true;
            return;
        }
        int anchor = Anchor();
        int length = _range.End - _range.Start;
        foreach (TextUnit unit in Ladder)
        {
            RangeWrapper around = UnitAround(anchor, unit);
            if (around.Start <= _range.Start && around.End >= _range.End && around.End - around.Start > length)
            {
                SetLadder(around.Start, around.End, anchor);
                return;
            }
        }
    }

    /// <summary>
    /// Extends the selection up to and including the next occurrence of a character after it, as <c>Selection.Extend</c> in Word with
    /// a character; the case counts. A selection that was not made by <see cref="Extend()"/> or this method is dropped first: the new one
    /// starts where it ended. Nothing changes if the character does not occur. Extend mode is left as it is.
    /// </summary>
    /// <param name="character">One character.</param>
    public void Extend(string character)
    {
        Guard.ThrowIfNull(character);
        if (character.Length != 1)
        {
            return;
        }
        int from = _range.End;
        int anchor = _ladderStart == _range.Start && _ladderEnd == _range.End ? _ladderAnchor : from;
        RangeWrapper rest = _range.Duplicate();
        rest.SetRange(from, _range.StoryLength);
        int index = rest.Text.IndexOf(character, StringComparison.Ordinal);
        if (index < 0)
        {
            return;
        }
        SetLadder(anchor, from + index + 1, anchor);
    }

    /// <summary>
    /// Makes the selection the largest unit around the place it was extended from (the story, a paragraph, a sentence, a word) that is
    /// smaller than it is, or an insertion point there when there is none, as <c>Selection.Shrink</c> in Word.
    /// </summary>
    public void Shrink()
    {
        int length = _range.End - _range.Start;
        if (length == 0)
        {
            return;
        }
        int anchor = Anchor();
        foreach (TextUnit unit in LadderDown)
        {
            RangeWrapper around = UnitAround(anchor, unit);
            if (around.End - around.Start < length)
            {
                SetLadder(around.Start, around.End, anchor);
                return;
            }
        }
        SetLadder(anchor, anchor, anchor);
    }

    /// <summary>
    /// Selects the next field after the insertion point or the start of the selection, as <c>Selection.NextField</c> in Word; the field
    /// that is selected already does not count. Returns null, and selects nothing, if there is none.
    /// </summary>
    public FieldWrapper NextField()
    {
        int start = _range.Start;
        int end = _range.End;
        FieldWrapper found = new RangeWrapper(StoryMap.FindRoot(_range.AnchorElement)).Fields
            .FirstOrDefault(f => f.Range.Start >= start && !(f.Range.Start == start && f.Range.End == end));
        if (found != null)
        {
            Set(found.Range);
        }
        return found;
    }

    /// <summary>
    /// Selects the last field that ends before the insertion point or the start of the selection, as <c>Selection.PreviousField</c> in
    /// Word. Returns null, and selects nothing, if there is none.
    /// </summary>
    public FieldWrapper PreviousField()
    {
        int start = _range.Start;
        FieldWrapper found = new RangeWrapper(StoryMap.FindRoot(_range.AnchorElement)).Fields
            .LastOrDefault(f => f.Range.End <= start);
        if (found != null)
        {
            Set(found.Range);
        }
        return found;
    }
}
