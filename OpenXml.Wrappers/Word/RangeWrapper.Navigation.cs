using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Moving a range to the start or end of a unit or until a character, comparing ranges, and measuring one. The rules, including the numbers
/// the methods return, are those that Word follows, measured in Word.
/// </summary>
public partial class RangeWrapper
{
    //Which end or ends of a range a move works on.
    internal enum StoryEnd
    {
        Both,
        Start,
        End
    }

    /// <summary>
    /// The value of <c>wdForward</c> in Word: a count for <see cref="MoveUntil"/> and the like that is as many characters as there are.
    /// A negative count of the same size (<c>wdBackward</c>) moves backwards.
    /// </summary>
    private const int Forward = 1073741823;

    /// <summary>
    /// Moves or extends the start of the range to the start of the unit it is in, as <c>Range.StartOf</c> in Word. A range that is already at
    /// the start of a unit is not extended. Returns the number of characters the start moved, which is negative, and one more than that for a
    /// <see cref="Movement.Move"/> of a range that is not collapsed, as in Word.
    /// </summary>
    /// <param name="unit">The unit, a word if it is left out.</param>
    /// <param name="extend"><see cref="Movement.Move"/> to collapse the range to the start of the unit, <see cref="Movement.Extend"/> to move only its start.</param>
    public int StartOf(TextUnit unit = TextUnit.Word, Movement extend = Movement.Move)
    {
        CheckMovement(extend);
        GetPositions(out StoryMap map, out int start, out int end);
        int target = unit == TextUnit.Story ? 0 : TextUnits.GetBoundaries(map.ToCharacters(), unit).LastOrDefault(b => b <= start);
        if (extend == Movement.Extend)
        {
            MakeSpan(map, target, end);
            return target - start;
        }
        MakeSpan(map, target, target);
        return target - start - (start < end ? 1 : 0);
    }

    /// <summary>
    /// Moves or extends the end of the range to the end of the unit it is in, as <c>Range.EndOf</c> in Word. A range that already ends at the
    /// end of a unit is not extended. Returns the number of characters the end moved, and one more than that for a <see cref="Movement.Move"/>
    /// of a range that is not collapsed, as in Word. A range that is moved cannot go after the last paragraph mark of the story.
    /// </summary>
    /// <param name="unit">The unit, a word if it is left out.</param>
    /// <param name="extend"><see cref="Movement.Move"/> to collapse the range to the end of the unit, <see cref="Movement.Extend"/> to move only its end.</param>
    public int EndOf(TextUnit unit = TextUnit.Word, Movement extend = Movement.Move)
    {
        CheckMovement(extend);
        GetPositions(out StoryMap map, out int start, out int end);

        //A range that is not collapsed ends in the unit that holds its last character, a collapsed one in the unit that holds its position.
        int after = start < end ? end - 1 : end;
        int target = unit == TextUnit.Story ? map.Length : TextUnits.GetBoundaries(map.ToCharacters(), unit).FirstOrDefault(b => b > after);
        if (target < end)
        {
            target = map.Length;
        }
        if (extend == Movement.Extend)
        {
            MakeSpan(map, start, target);
            return target - end;
        }
        target = Math.Min(target, LastInsertionPoint(map));
        MakeSpan(map, target, target);
        return target - end + (start < end ? 1 : 0);
    }

    /// <summary>
    /// The unit that follows the range, as <c>Range.Next</c> in Word: the first unit that starts at or after the end of the range (after the
    /// position of a collapsed range), or the one <paramref name="count"/> units on. A negative count is <see cref="Previous"/>. Returns null
    /// if there is no such unit.
    /// </summary>
    /// <param name="unit">The unit, a character if it is left out.</param>
    /// <param name="count">The number of units to count, one if it is left out.</param>
    /// <exception cref="ArgumentException">The unit is <see cref="TextUnit.Story"/>.</exception>
    public RangeWrapper Next(TextUnit unit = TextUnit.Character, int count = 1)
    {
        CheckCountable(unit);
        if (count < 0)
        {
            return Previous(unit, -count);
        }
        if (count == 0)
        {
            return null;
        }

        GetPositions(out StoryMap map, out int start, out int end);
        List<int> boundaries = TextUnits.GetBoundaries(map.ToCharacters(), unit);
        int after = start < end ? end - 1 : end;
        int first = boundaries.FindIndex(b => b > after);
        int index = first < 0 ? -1 : first + count - 1;
        if (index < 0 || index + 1 >= boundaries.Count || boundaries[index] >= map.Length)
        {
            return null;
        }
        return FromMap(map, boundaries[index], boundaries[index + 1]);
    }

    /// <summary>
    /// The unit that comes before the range, as <c>Range.Previous</c> in Word: the unit that ends where the unit holding the start of the
    /// range begins, or the one <paramref name="count"/> units back. A negative count is <see cref="Next"/>. Returns null if there is no such
    /// unit.
    /// </summary>
    /// <param name="unit">The unit, a character if it is left out.</param>
    /// <param name="count">The number of units to count, one if it is left out.</param>
    /// <exception cref="ArgumentException">The unit is <see cref="TextUnit.Story"/>.</exception>
    public RangeWrapper Previous(TextUnit unit = TextUnit.Character, int count = 1)
    {
        CheckCountable(unit);
        if (count < 0)
        {
            return Next(unit, -count);
        }
        if (count == 0)
        {
            return null;
        }

        GetPositions(out StoryMap map, out int start, out _);
        List<int> boundaries = TextUnits.GetBoundaries(map.ToCharacters(), unit);
        int holding = boundaries.FindLastIndex(b => b <= start);
        int index = holding - count;
        if (holding < 0 || index < 0)
        {
            return null;
        }
        return FromMap(map, boundaries[index], boundaries[index + 1]);
    }

    /// <summary>
    /// Moves the range, collapsed, until one of the characters is found, as <c>Range.MoveUntil</c> in Word. It looks forward from the end of
    /// the range (up to <paramref name="count"/> characters) or, when the count is negative, backward from the start, and stops before the
    /// character going forward and after it going back. The match is case sensitive. Returns the number of characters it moved, one more
    /// forward and one less back, and 0 (leaving the range alone) if none of the characters is found.
    /// </summary>
    /// <param name="characters">The characters to stop at.</param>
    /// <param name="count">The most characters to move, forward if positive; as many as there are if it is left out.</param>
    public int MoveUntil(string characters, int count = Forward)
    {
        return MoveByCharacters(StoryEnd.Both, characters, count, true);
    }

    /// <summary>
    /// Moves the range, collapsed, past the characters that are in the set, as <c>Range.MoveWhile</c> in Word. Returns the number of characters
    /// it moved (negative backward), and 0 leaving the range alone if the first character is not one of them. See <see cref="MoveUntil"/>.
    /// </summary>
    /// <param name="characters">The characters to move over.</param>
    /// <param name="count">The most characters to move, forward if positive; as many as there are if it is left out.</param>
    public int MoveWhile(string characters, int count = Forward)
    {
        return MoveByCharacters(StoryEnd.Both, characters, count, false);
    }

    /// <summary>
    /// Moves the start of the range until one of the characters is found, as <c>Range.MoveStartUntil</c> in Word. The range is collapsed
    /// where it is moved to if that is past its end. See <see cref="MoveUntil"/>.
    /// </summary>
    /// <param name="characters">The characters to stop at.</param>
    /// <param name="count">The most characters to move, forward if positive; as many as there are if it is left out.</param>
    public int MoveStartUntil(string characters, int count = Forward)
    {
        return MoveByCharacters(StoryEnd.Start, characters, count, true);
    }

    /// <summary>
    /// Moves the start of the range past the characters that are in the set, as <c>Range.MoveStartWhile</c> in Word. See
    /// <see cref="MoveWhile"/>.
    /// </summary>
    /// <param name="characters">The characters to move over.</param>
    /// <param name="count">The most characters to move, forward if positive; as many as there are if it is left out.</param>
    public int MoveStartWhile(string characters, int count = Forward)
    {
        return MoveByCharacters(StoryEnd.Start, characters, count, false);
    }

    /// <summary>
    /// Moves the end of the range until one of the characters is found, as <c>Range.MoveEndUntil</c> in Word. The range is collapsed where it
    /// is moved to if that is before its start. See <see cref="MoveUntil"/>.
    /// </summary>
    /// <param name="characters">The characters to stop at.</param>
    /// <param name="count">The most characters to move, forward if positive; as many as there are if it is left out.</param>
    public int MoveEndUntil(string characters, int count = Forward)
    {
        return MoveByCharacters(StoryEnd.End, characters, count, true);
    }

    /// <summary>
    /// Moves the end of the range past the characters that are in the set, as <c>Range.MoveEndWhile</c> in Word. See <see cref="MoveWhile"/>.
    /// </summary>
    /// <param name="characters">The characters to move over.</param>
    /// <param name="count">The most characters to move, forward if positive; as many as there are if it is left out.</param>
    public int MoveEndWhile(string characters, int count = Forward)
    {
        return MoveByCharacters(StoryEnd.End, characters, count, false);
    }

    /// <summary>
    /// True if the other range is in the same story and starts and ends at the same places, as <c>Range.IsEqual</c> in Word.
    /// </summary>
    public bool IsEqual(RangeWrapper range)
    {
        Guard.ThrowIfNull(range, nameof(range));
        if (!InStory(range))
        {
            return false;
        }
        GetStoryPositions(out int start, out int end);
        range.GetStoryPositions(out int otherStart, out int otherEnd);
        return start == otherStart && end == otherEnd;
    }

    /// <summary>
    /// True if the range is inside the other range (or the same as it), in the same story, as <c>Range.InRange</c> in Word.
    /// </summary>
    public bool InRange(RangeWrapper range)
    {
        Guard.ThrowIfNull(range, nameof(range));
        if (!InStory(range))
        {
            return false;
        }
        GetStoryPositions(out int start, out int end);
        range.GetStoryPositions(out int otherStart, out int otherEnd);
        return start >= otherStart && end <= otherEnd;
    }

    /// <summary>
    /// True if the other range is in the same story (the body, a header, a footer or a comment), as <c>Range.InStory</c> in Word.
    /// </summary>
    public bool InStory(RangeWrapper range)
    {
        Guard.ThrowIfNull(range, nameof(range));
        return ReferenceEquals(StoryMap.FindRoot(AnchorElement), StoryMap.FindRoot(range.AnchorElement));
    }

    /// <summary>
    /// Makes the range the whole of its story, up to and including the last paragraph mark, as <c>Range.WholeStory</c> in Word.
    /// </summary>
    public void WholeStory()
    {
        GetPositions(out StoryMap map, out _, out _);
        MakeSpan(map, 0, map.Length);
    }

    /// <summary>
    /// Counts something in the range, as <c>Range.ComputeStatistics</c> in Word. Words are what is between spaces, so a dash on its own is
    /// one; <see cref="Statistic.Characters"/> does not count spaces and <see cref="Statistic.CharactersWithSpaces"/> does; neither counts
    /// paragraph marks. <see cref="Statistic.Paragraphs"/> counts the paragraph marks that follow some text of the range, and not the
    /// last paragraph mark of the story, so a range of the whole body of a document of three paragraphs has two.
    /// </summary>
    /// <remarks>
    /// Word gives 0 for everything in a range that includes the end of a table cell, which its own documentation calls inaccurate; the cell
    /// is counted here.
    /// </remarks>
    /// <exception cref="NotSupportedException">The statistic depends on the layout of the pages (<see cref="Statistic.Lines"/> and <see cref="Statistic.Pages"/>).</exception>
    public int ComputeStatistics(Statistic statistic)
    {
        if (!EnumCompat.IsDefined(statistic))
        {
            throw new ArgumentOutOfRangeException(nameof(statistic));
        }
        if (statistic == Statistic.Lines || statistic == Statistic.Pages)
        {
            throw new NotSupportedException("Lines and pages depend on the layout of the pages, which the library does not work out.");
        }

        GetPositions(out StoryMap map, out int start, out int end);
        string text = map.GetText(start, end);
        if (statistic == Statistic.Words)
        {
            return text.Split(new[] { ' ', '\t', '\r', '\n', '\a', '\v', '\f', ' ', '\0' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }
        if (statistic == Statistic.Characters)
        {
            return text.Count(c => !char.IsWhiteSpace(c) && !char.IsControl(c));
        }
        if (statistic == Statistic.CharactersWithSpaces)
        {
            return text.Count(c => !char.IsControl(c) || c == '\t');
        }

        int paragraphs = 0;
        bool hasText = false;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] == '\r')
            {
                if (hasText && start + i < map.Length - 1)
                {
                    paragraphs++;
                }
                hasText = false;
            }
            else if (text[i] != '\a' && text[i] != '\0')
            {
                hasText = true;
            }
        }
        return paragraphs;
    }

    /// <summary>
    /// Replaces the range with a paragraph mark, as <c>Range.InsertParagraph</c> in Word: the text is split there, and the range is left as
    /// the new mark.
    /// </summary>
    public void InsertParagraph()
    {
        Text = "\r";
    }

    //Moves one or both ends over characters until one is found (until) or until one is not in the set (while).
    private int MoveByCharacters(StoryEnd ends, string characters, int count, bool until)
    {
        Guard.ThrowIfNull(characters, nameof(characters));
        if (count == 0)
        {
            return 0;
        }

        GetPositions(out StoryMap map, out int start, out int end);
        char[] text = map.ToCharacters();
        char[] set = characters.ToCharArray();
        bool forward = count > 0;

        //Both ends move from the end going forward and from the start going back; one end moves from itself.
        long origin = ends == StoryEnd.Start ? start : ends == StoryEnd.End ? end : forward ? end : start;

        //The last paragraph mark of the story cannot be passed forward; the characters looked at are those between the origin and the limit.
        long last = Math.Min(origin + count, LastInsertionPoint(map));
        long first = Math.Max(origin + count, 0);
        int moved = 0;
        int position = (int)origin;
        bool found = false;
        if (forward)
        {
            for (long i = origin; i < last; i++)
            {
                bool inSet = text[i] != '\0' && Array.IndexOf(set, text[i]) >= 0;
                if (inSet == until)
                {
                    found = until;
                    break;
                }
                moved++;
            }
            position = (int)(origin + moved);
        }
        else
        {
            for (long i = origin - 1; i >= first; i--)
            {
                bool inSet = text[i] != '\0' && Array.IndexOf(set, text[i]) >= 0;
                if (inSet == until)
                {
                    found = until;
                    break;
                }
                moved++;
            }
            position = (int)(origin - moved);
        }

        int result;
        if (until)
        {
            if (!found)
            {
                return 0;
            }

            //Forward, the range is before the character it found; back, after it.
            result = forward ? moved + 1 : -moved - 1;
        }
        else
        {
            if (moved == 0)
            {
                return 0;
            }
            result = forward ? moved : -moved;
        }

        if (ends == StoryEnd.Both)
        {
            MakeSpan(map, position, position);
        }
        else if (ends == StoryEnd.Start)
        {
            MakeSpan(map, position, Math.Max(position, end));
        }
        else
        {
            MakeSpan(map, Math.Min(position, start), position);
        }
        return result;
    }

    //Moves by the one unit that a story is made of. It cannot go after the last paragraph mark unless it is the end that moves.
    private int MoveByStory(StoryEnd ends, int count)
    {
        if (count == 0)
        {
            return 0;
        }

        GetPositions(out StoryMap map, out int start, out int end);
        bool forward = count > 0;
        int stop = LastInsertionPoint(map);
        if (ends == StoryEnd.Both)
        {
            int target = forward ? stop : 0;
            if (start == end && start == target)
            {
                return 0;
            }
            MakeSpan(map, target, target);
        }
        else if (ends == StoryEnd.Start)
        {
            int target = forward ? stop : 0;
            if (start == target)
            {
                return 0;
            }
            MakeSpan(map, target, Math.Max(target, end));
        }
        else
        {
            int target = forward ? map.Length : 0;
            if (end == target)
            {
                return 0;
            }
            MakeSpan(map, Math.Min(start, target), target);
        }
        return forward ? 1 : -1;
    }

    //The last place a collapsed range can be in a story: before the last paragraph mark.
    private static int LastInsertionPoint(StoryMap map)
    {
        return Math.Max(0, map.Length - 1);
    }

    private static void CheckMovement(Movement extend)
    {
        if (extend != Movement.Move && extend != Movement.Extend)
        {
            throw new ArgumentOutOfRangeException(nameof(extend));
        }
    }

    private static void CheckCountable(TextUnit unit)
    {
        if (unit == TextUnit.Story)
        {
            throw new ArgumentException("A story is one unit; use WholeStory.", nameof(unit));
        }
        if (unit != TextUnit.Character && unit != TextUnit.Word && unit != TextUnit.Sentence && unit != TextUnit.Paragraph)
        {
            throw new ArgumentOutOfRangeException(nameof(unit));
        }
    }
}
