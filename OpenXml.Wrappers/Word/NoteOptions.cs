using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>How the marks of footnotes and endnotes are written, as <c>WdNoteNumberStyle</c> in Word (with the same values).</summary>
public enum NoteNumberStyle
{
    /// <summary>1, 2, 3.</summary>
    Arabic = 0,

    /// <summary>I, II, III.</summary>
    UppercaseRoman = 1,

    /// <summary>i, ii, iii.</summary>
    LowercaseRoman = 2,

    /// <summary>A, B, C.</summary>
    UppercaseLetter = 3,

    /// <summary>a, b, c.</summary>
    LowercaseLetter = 4,

    /// <summary>*, †, ‡, §.</summary>
    Symbol = 9
}

/// <summary>When the numbering of notes starts again, as <c>WdNumberingRule</c> in Word (with the same values).</summary>
public enum NumberingRule
{
    /// <summary>It does not start again.</summary>
    Continuous = 0,

    /// <summary>In each section.</summary>
    RestartSection = 1,

    /// <summary>On each page.</summary>
    RestartPage = 2
}

/// <summary>Where footnotes go, as <c>WdFootnoteLocation</c> in Word (with the same values).</summary>
public enum FootnoteLocation
{
    /// <summary>At the bottom of the page.</summary>
    BottomOfPage = 0,

    /// <summary>Just below the text.</summary>
    BeneathText = 1
}

/// <summary>Where endnotes go, as <c>WdEndnoteLocation</c> in Word (with the same values).</summary>
public enum EndnoteLocation
{
    /// <summary>At the end of each section.</summary>
    EndOfSection = 0,

    /// <summary>At the end of the document.</summary>
    EndOfDocument = 1
}

/// <summary>
/// The options of footnotes or endnotes kept in one place: the settings of the document, or the properties of a section. Both are the same
/// elements under different names.
/// </summary>
internal sealed class NoteOptionsStore
{
    private readonly bool _endnote;
    private readonly Func<bool, OpenXmlCompositeElement> _get;

    public NoteOptionsStore(bool endnote, Func<bool, OpenXmlCompositeElement> get)
    {
        _endnote = endnote;
        _get = get;
    }

    private T Find<T>() where T : OpenXmlElement
    {
        return _get(false)?.GetFirstChild<T>();
    }

    private void Put(OpenXmlElement element, Type type)
    {
        OpenXmlCompositeElement container = _get(true);
        foreach (OpenXmlElement existing in container.ChildElements.Where(e => e.GetType() == type).ToList())
        {
            existing.Remove();
        }
        if (element != null)
        {
            container.AddChild(element);
        }
    }

    public bool TryGetStyle(out NoteNumberStyle style)
    {
        NumberFormatValues? format = Find<NumberingFormat>()?.Val?.Value;
        style = format == NumberFormatValues.UpperRoman ? NoteNumberStyle.UppercaseRoman
            : format == NumberFormatValues.LowerRoman ? NoteNumberStyle.LowercaseRoman
            : format == NumberFormatValues.UpperLetter ? NoteNumberStyle.UppercaseLetter
            : format == NumberFormatValues.LowerLetter ? NoteNumberStyle.LowercaseLetter
            : format == NumberFormatValues.Chicago ? NoteNumberStyle.Symbol : NoteNumberStyle.Arabic;
        return format != null;
    }

    public void SetStyle(NoteNumberStyle style)
    {
        Put(new NumberingFormat()
        {
            Val = style == NoteNumberStyle.UppercaseRoman ? NumberFormatValues.UpperRoman
            : style == NoteNumberStyle.LowercaseRoman ? NumberFormatValues.LowerRoman
            : style == NoteNumberStyle.UppercaseLetter ? NumberFormatValues.UpperLetter
            : style == NoteNumberStyle.LowercaseLetter ? NumberFormatValues.LowerLetter
            : style == NoteNumberStyle.Symbol ? NumberFormatValues.Chicago : NumberFormatValues.Decimal
        }, typeof(NumberingFormat));
    }

    public bool TryGetStart(out int start)
    {
        start = Find<NumberingStart>()?.Val?.Value ?? 1;
        return Find<NumberingStart>() != null;
    }

    public void SetStart(int start)
    {
        Put(start == 1 ? null : new NumberingStart() { Val = (ushort)start }, typeof(NumberingStart));
    }

    public bool TryGetRule(out NumberingRule rule)
    {
        RestartNumberValues? restart = Find<NumberingRestart>()?.Val?.Value;
        rule = restart == RestartNumberValues.EachSection ? NumberingRule.RestartSection : restart == RestartNumberValues.EachPage ? NumberingRule.RestartPage : NumberingRule.Continuous;
        return restart != null;
    }

    public void SetRule(NumberingRule rule)
    {
        Put(rule == NumberingRule.Continuous ? null : new NumberingRestart() { Val = rule == NumberingRule.RestartSection ? RestartNumberValues.EachSection : RestartNumberValues.EachPage }, typeof(NumberingRestart));
    }

    public bool TryGetLocation(out int location)
    {
        if (_endnote)
        {
            EndnotePositionValues? endnotePosition = Find<EndnotePosition>()?.Val?.Value;
            location = endnotePosition == EndnotePositionValues.SectionEnd ? 0 : 1;
            return endnotePosition != null;
        }
        FootnotePositionValues? position = Find<FootnotePosition>()?.Val?.Value;
        location = position == FootnotePositionValues.BeneathText ? 1 : 0;
        return position != null;
    }

    public void SetLocation(int location)
    {
        if (_endnote)
        {
            Put(location == 0 ? new EndnotePosition() { Val = EndnotePositionValues.SectionEnd } : null, typeof(EndnotePosition));
            return;
        }
        Put(location == 1 ? new FootnotePosition() { Val = FootnotePositionValues.BeneathText } : null, typeof(FootnotePosition));
    }
}

/// <summary>
/// The options of footnotes or endnotes for the document as a whole: kept in its settings and in the properties of every section, as Word
/// keeps them when they are set on <c>Footnotes</c> or <c>Endnotes</c>.
/// </summary>
internal static class NoteOptionsOps
{
    private static OpenXmlCompositeElement GetDocumentContainer(MainDocumentPart main, bool endnote, bool create)
    {
        Settings settings = main.DocumentSettingsPart?.Settings;
        if (settings == null)
        {
            return null;
        }
        if (endnote)
        {
            return settings.GetFirstChild<EndnoteDocumentWideProperties>();
        }
        return settings.GetFirstChild<FootnoteDocumentWideProperties>();
    }

    internal static NoteOptionsStore Document(MainDocumentPart main, bool endnote)
    {
        return new NoteOptionsStore(endnote, create =>
        {
            if (create)
            {
                NoteOps.EnsureParts(main);
            }
            return GetDocumentContainer(main, endnote, create);
        });
    }

    internal static NoteOptionsStore Section(SectionWrapper section, bool endnote)
    {
        return new NoteOptionsStore(endnote, create =>
        {
            SectionProperties properties = section.GetSectionProperties(create);
            if (properties == null)
            {
                return null;
            }
            if (endnote)
            {
                EndnoteProperties existing = properties.GetFirstChild<EndnoteProperties>();
                if (existing == null && create)
                {
                    existing = new EndnoteProperties();
                    properties.AddChild(existing);
                }
                return existing;
            }
            FootnoteProperties note = properties.GetFirstChild<FootnoteProperties>();
            if (note == null && create)
            {
                note = new FootnoteProperties();
                properties.AddChild(note);
            }
            return note;
        });
    }
}
