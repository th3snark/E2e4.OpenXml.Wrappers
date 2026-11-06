using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Reads and writes the options of footnotes or endnotes on a list of places in turn: the first that has the option answers, and a change is
/// made in every place.
/// </summary>
internal sealed class NoteOptionsView
{
    private readonly bool _endnote;
    private readonly IReadOnlyList<NoteOptionsStore> _readFrom;
    private readonly IReadOnlyList<NoteOptionsStore> _writeTo;

    public NoteOptionsView(bool endnote, IReadOnlyList<NoteOptionsStore> readFrom, IReadOnlyList<NoteOptionsStore> writeTo)
    {
        _endnote = endnote;
        _readFrom = readFrom;
        _writeTo = writeTo;
    }

    public NoteNumberStyle NumberStyle
    {
        get
        {
            foreach (NoteOptionsStore store in _readFrom)
            {
                if (store.TryGetStyle(out NoteNumberStyle style))
                {
                    return style;
                }
            }
            //Endnotes are numbered in lowercase roman numbers unless the document says otherwise, as in Word.
            return _endnote ? NoteNumberStyle.LowercaseRoman : NoteNumberStyle.Arabic;
        }
    }

    public int StartingNumber
    {
        get
        {
            foreach (NoteOptionsStore store in _readFrom)
            {
                if (store.TryGetStart(out int start))
                {
                    return start;
                }
            }
            return 1;
        }
    }

    public NumberingRule NumberingRule
    {
        get
        {
            foreach (NoteOptionsStore store in _readFrom)
            {
                if (store.TryGetRule(out NumberingRule rule))
                {
                    return rule;
                }
            }
            return NumberingRule.Continuous;
        }
    }

    public int Location
    {
        get
        {
            foreach (NoteOptionsStore store in _readFrom)
            {
                if (store.TryGetLocation(out int location))
                {
                    return location;
                }
            }
            return _endnote ? 1 : 0;
        }
    }

    public void SetNumberStyle(NoteNumberStyle value)
    {
        if (!EnumCompat.IsDefined(value))
        {
            throw new ArgumentOutOfRangeException(nameof(value));
        }
        //The default of the kind of note is not written, as Word does not write it for the footnotes it leaves at the default.
        foreach (NoteOptionsStore store in _writeTo)
        {
            store.SetStyle(value);
        }
    }

    public void SetStartingNumber(int value)
    {
        if (value < 1 || value > 16383)
        {
            throw new ArgumentOutOfRangeException(nameof(value), "The number must be between 1 and 16383.");
        }
        if (value != 1 && NumberingRule != NumberingRule.Continuous)
        {
            throw new ArgumentException("Notes must start at 1 if the numbering is not continuous.", nameof(value));
        }
        foreach (NoteOptionsStore store in _writeTo)
        {
            store.SetStart(value);
        }
    }

    public void SetNumberingRule(NumberingRule value)
    {
        if (!EnumCompat.IsDefined(value))
        {
            throw new ArgumentOutOfRangeException(nameof(value));
        }
        if (value != NumberingRule.Continuous && StartingNumber != 1)
        {
            throw new ArgumentException("Notes must start at 1 if the numbering is not continuous.", nameof(value));
        }
        foreach (NoteOptionsStore store in _writeTo)
        {
            store.SetRule(value);
        }
    }

    public void SetLocation(int value, bool documentOnly)
    {
        foreach (NoteOptionsStore store in documentOnly ? _writeTo.Take(1) : _writeTo)
        {
            store.SetLocation(value);
        }
    }
}

internal static class NoteOptionsViews
{
    //The options of the document: its settings, and every section.
    internal static NoteOptionsView ForDocument(WordprocessingDocumentWrapper document, bool endnote)
    {
        MainDocumentPart main = document.WordprocessingDocument.MainDocumentPart;
        List<NoteOptionsStore> write = new List<NoteOptionsStore> { NoteOptionsOps.Document(main, endnote) };
        write.AddRange(document.Sections.Select(s => NoteOptionsOps.Section(s, endnote)));
        return new NoteOptionsView(endnote, new[] { write[0] }, write);
    }

    //The options of the sections a range is in: their own, and the document's where they have none.
    internal static NoteOptionsView ForRange(RangeWrapper range, bool endnote)
    {
        WordprocessingDocumentWrapper document = range.Document;
        List<SectionWrapper> sections = range.Sections.ToList();
        List<NoteOptionsStore> sectionStores = sections.Select(s => NoteOptionsOps.Section(s, endnote)).ToList();
        List<NoteOptionsStore> read = new List<NoteOptionsStore>();
        if (sectionStores.Count > 0)
        {
            read.Add(sectionStores[0]);
        }
        read.Add(NoteOptionsOps.Document(document.WordprocessingDocument.MainDocumentPart, endnote));
        return new NoteOptionsView(endnote, read, sectionStores);
    }
}

public partial class FootnoteWrapperCollection
{
    private NoteOptionsView Options
    {
        get { return NoteOptionsViews.ForDocument(_range.Document, false); }
    }

    /// <summary>Where the footnotes go, as <c>Footnotes.Location</c> in Word (bottom of the page by default).</summary>
    public FootnoteLocation Location
    {
        get { return (FootnoteLocation)Options.Location; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Options.SetLocation((int)value, false);
        }
    }

    /// <summary>How the marks are written, as <c>Footnotes.NumberStyle</c> in Word.</summary>
    public NoteNumberStyle NumberStyle
    {
        get { return Options.NumberStyle; }
        set { Options.SetNumberStyle(value); }
    }

    /// <summary>When the numbering starts again, as <c>Footnotes.NumberingRule</c> in Word.</summary>
    /// <exception cref="ArgumentException">The numbering starts at a number other than 1 and so must be continuous.</exception>
    public NumberingRule NumberingRule
    {
        get { return Options.NumberingRule; }
        set { Options.SetNumberingRule(value); }
    }

    /// <summary>The number of the first footnote, as <c>Footnotes.StartingNumber</c> in Word.</summary>
    /// <exception cref="ArgumentException">The number is not 1 and the numbering starts again.</exception>
    public int StartingNumber
    {
        get { return Options.StartingNumber; }
        set { Options.SetStartingNumber(value); }
    }

    /// <summary>The line above the footnotes, as <c>Footnotes.Separator</c> in Word.</summary>
    public RangeWrapper Separator
    {
        get { return NoteOps.GetSeparator(StyleHelper.GetMainDocumentPart(_range.AnchorElement), false, FootnoteEndnoteValues.Separator); }
    }

    /// <summary>The line above the footnotes that carry on from the page before, as <c>Footnotes.ContinuationSeparator</c> in Word.</summary>
    public RangeWrapper ContinuationSeparator
    {
        get { return NoteOps.GetSeparator(StyleHelper.GetMainDocumentPart(_range.AnchorElement), false, FootnoteEndnoteValues.ContinuationSeparator); }
    }

    /// <summary>
    /// Makes all the footnotes of the document endnotes, as <c>Footnotes.Convert</c> in Word.
    /// </summary>
    public void Convert()
    {
        NoteOps.Convert(StyleHelper.GetMainDocumentPart(_range.AnchorElement), false);
    }

    /// <summary>Makes the footnotes endnotes and the endnotes footnotes, as <c>Footnotes.SwapWithEndnotes</c> in Word.</summary>
    public void SwapWithEndnotes()
    {
        NoteOps.Swap(StyleHelper.GetMainDocumentPart(_range.AnchorElement));
    }
}

public partial class EndnoteWrapperCollection
{
    private NoteOptionsView Options
    {
        get { return NoteOptionsViews.ForDocument(_range.Document, true); }
    }

    /// <summary>Where the endnotes go, as <c>Endnotes.Location</c> in Word (the end of the document by default).</summary>
    public EndnoteLocation Location
    {
        get { return (EndnoteLocation)Options.Location; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Options.SetLocation((int)value, true);
        }
    }

    /// <summary>How the marks are written, as <c>Endnotes.NumberStyle</c> in Word (lowercase roman numbers by default).</summary>
    public NoteNumberStyle NumberStyle
    {
        get { return Options.NumberStyle; }
        set { Options.SetNumberStyle(value); }
    }

    /// <summary>When the numbering starts again, as <c>Endnotes.NumberingRule</c> in Word.</summary>
    /// <exception cref="ArgumentException">The numbering starts at a number other than 1 and so must be continuous.</exception>
    public NumberingRule NumberingRule
    {
        get { return Options.NumberingRule; }
        set { Options.SetNumberingRule(value); }
    }

    /// <summary>The number of the first endnote, as <c>Endnotes.StartingNumber</c> in Word.</summary>
    /// <exception cref="ArgumentException">The number is not 1 and the numbering starts again.</exception>
    public int StartingNumber
    {
        get { return Options.StartingNumber; }
        set { Options.SetStartingNumber(value); }
    }

    /// <summary>The line above the endnotes, as <c>Endnotes.Separator</c> in Word.</summary>
    public RangeWrapper Separator
    {
        get { return NoteOps.GetSeparator(StyleHelper.GetMainDocumentPart(_range.AnchorElement), true, FootnoteEndnoteValues.Separator); }
    }

    /// <summary>The line above the endnotes that carry on from the page before, as <c>Endnotes.ContinuationSeparator</c> in Word.</summary>
    public RangeWrapper ContinuationSeparator
    {
        get { return NoteOps.GetSeparator(StyleHelper.GetMainDocumentPart(_range.AnchorElement), true, FootnoteEndnoteValues.ContinuationSeparator); }
    }

    /// <summary>Makes all the endnotes of the document footnotes, as <c>Endnotes.Convert</c> in Word.</summary>
    public void Convert()
    {
        NoteOps.Convert(StyleHelper.GetMainDocumentPart(_range.AnchorElement), true);
    }

    /// <summary>Makes the endnotes footnotes and the footnotes endnotes, as <c>Endnotes.SwapWithFootnotes</c> in Word.</summary>
    public void SwapWithFootnotes()
    {
        NoteOps.Swap(StyleHelper.GetMainDocumentPart(_range.AnchorElement));
    }
}

/// <summary>
/// The options of the footnotes of the sections of a range, as <c>FootnoteOptions</c> in Word. They are the options of the sections, with
/// those of the document for what a section has not set.
/// </summary>
public class FootnoteOptions
{
    private readonly RangeWrapper _range;

    internal FootnoteOptions(RangeWrapper range)
    {
        _range = range;
    }

    private NoteOptionsView Options
    {
        get { return NoteOptionsViews.ForRange(_range, false); }
    }

    /// <summary>The range the options are of.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    /// <summary>Where the footnotes go, as <c>FootnoteOptions.Location</c> in Word.</summary>
    public FootnoteLocation Location
    {
        get { return (FootnoteLocation)Options.Location; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Options.SetLocation((int)value, false);
        }
    }

    /// <summary>How the marks are written, as <c>FootnoteOptions.NumberStyle</c> in Word.</summary>
    public NoteNumberStyle NumberStyle
    {
        get { return Options.NumberStyle; }
        set { Options.SetNumberStyle(value); }
    }

    /// <summary>When the numbering starts again, as <c>FootnoteOptions.NumberingRule</c> in Word.</summary>
    public NumberingRule NumberingRule
    {
        get { return Options.NumberingRule; }
        set { Options.SetNumberingRule(value); }
    }

    /// <summary>The number of the first footnote, as <c>FootnoteOptions.StartingNumber</c> in Word.</summary>
    public int StartingNumber
    {
        get { return Options.StartingNumber; }
        set { Options.SetStartingNumber(value); }
    }
}

/// <summary>
/// The options of the endnotes of the sections of a range, as <c>EndnoteOptions</c> in Word. The location is the document's: it is not kept
/// in a section.
/// </summary>
public class EndnoteOptions
{
    private readonly RangeWrapper _range;

    internal EndnoteOptions(RangeWrapper range)
    {
        _range = range;
    }

    private NoteOptionsView Options
    {
        get { return NoteOptionsViews.ForRange(_range, true); }
    }

    /// <summary>The range the options are of.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    /// <summary>Where the endnotes go, as <c>EndnoteOptions.Location</c> in Word; a setting of the whole document.</summary>
    public EndnoteLocation Location
    {
        get { return (EndnoteLocation)Options.Location; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            NoteOptionsViews.ForDocument(_range.Document, true).SetLocation((int)value, true);
        }
    }

    /// <summary>How the marks are written, as <c>EndnoteOptions.NumberStyle</c> in Word.</summary>
    public NoteNumberStyle NumberStyle
    {
        get { return Options.NumberStyle; }
        set { Options.SetNumberStyle(value); }
    }

    /// <summary>When the numbering starts again, as <c>EndnoteOptions.NumberingRule</c> in Word.</summary>
    public NumberingRule NumberingRule
    {
        get { return Options.NumberingRule; }
        set { Options.SetNumberingRule(value); }
    }

    /// <summary>The number of the first endnote, as <c>EndnoteOptions.StartingNumber</c> in Word.</summary>
    public int StartingNumber
    {
        get { return Options.StartingNumber; }
        set { Options.SetStartingNumber(value); }
    }
}
