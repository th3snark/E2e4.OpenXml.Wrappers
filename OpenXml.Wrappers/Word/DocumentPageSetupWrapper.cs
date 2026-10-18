using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The page setup of the whole document, as <c>Document.PageSetup</c> in Word: reading gives the value of the first section, and setting
/// sets it in every section. The things that belong to a section in a more detailed way (columns, line numbers and page borders) are set on
/// <c>Sections[i].PageSetup</c>.
/// </summary>
public class DocumentPageSetupWrapper
{
    private readonly SectionWrapperCollection _sections;

    internal DocumentPageSetupWrapper(SectionWrapperCollection sections)
    {
        _sections = sections;
    }

    private PageSetupWrapper First
    {
        get { return _sections.First.PageSetup; }
    }

    private void Apply(Action<PageSetupWrapper> change)
    {
        foreach (SectionWrapper section in _sections)
        {
            change(section.PageSetup);
        }
    }

    /// <summary>The width of the pages in points.</summary>
    public double PageWidth
    {
        get { return First.PageWidth; }
        set { Apply(p => p.PageWidth = value); }
    }

    /// <summary>The height of the pages in points.</summary>
    public double PageHeight
    {
        get { return First.PageHeight; }
        set { Apply(p => p.PageHeight = value); }
    }

    /// <summary>The orientation of the pages. Setting it swaps the width and height of the sections that change.</summary>
    public PageOrientation Orientation
    {
        get { return First.Orientation; }
        set { Apply(p => p.Orientation = value); }
    }

    /// <summary>The paper size.</summary>
    public PaperSize PaperSize
    {
        get { return First.PaperSize; }
        set { Apply(p => p.PaperSize = value); }
    }

    /// <summary>The top margin in points.</summary>
    public double TopMargin
    {
        get { return First.TopMargin; }
        set { Apply(p => p.TopMargin = value); }
    }

    /// <summary>The bottom margin in points.</summary>
    public double BottomMargin
    {
        get { return First.BottomMargin; }
        set { Apply(p => p.BottomMargin = value); }
    }

    /// <summary>The left margin in points.</summary>
    public double LeftMargin
    {
        get { return First.LeftMargin; }
        set { Apply(p => p.LeftMargin = value); }
    }

    /// <summary>The right margin in points.</summary>
    public double RightMargin
    {
        get { return First.RightMargin; }
        set { Apply(p => p.RightMargin = value); }
    }

    /// <summary>The gutter in points.</summary>
    public double Gutter
    {
        get { return First.Gutter; }
        set { Apply(p => p.Gutter = value); }
    }

    /// <summary>The distance in points from the top of the page to the header.</summary>
    public double HeaderDistance
    {
        get { return First.HeaderDistance; }
        set { Apply(p => p.HeaderDistance = value); }
    }

    /// <summary>The distance in points from the bottom of the page to the footer.</summary>
    public double FooterDistance
    {
        get { return First.FooterDistance; }
        set { Apply(p => p.FooterDistance = value); }
    }

    /// <summary>Where the sections start.</summary>
    public SectionStart SectionStart
    {
        get { return First.SectionStart; }
        set { Apply(p => p.SectionStart = value); }
    }

    /// <summary>Whether the first page has a header and footer of its own.</summary>
    public bool DifferentFirstPageHeaderFooter
    {
        get { return First.DifferentFirstPageHeaderFooter; }
        set { Apply(p => p.DifferentFirstPageHeaderFooter = value); }
    }

    /// <summary>Whether odd and even pages have different headers and footers (a setting of the document).</summary>
    public bool OddAndEvenPagesHeaderFooter
    {
        get { return First.OddAndEvenPagesHeaderFooter; }
        set { First.OddAndEvenPagesHeaderFooter = value; }
    }

    /// <summary>How the text is lined up between the top and bottom margins.</summary>
    public PageVerticalAlignment VerticalAlignment
    {
        get { return First.VerticalAlignment; }
        set { Apply(p => p.VerticalAlignment = value); }
    }

    /// <summary>Whether the endnotes of the sections are left out.</summary>
    public bool SuppressEndnotes
    {
        get { return First.SuppressEndnotes; }
        set { Apply(p => p.SuppressEndnotes = value); }
    }

    /// <summary>Whether the margins are mirrored on facing pages (a setting of the document).</summary>
    public bool MirrorMargins
    {
        get { return First.MirrorMargins; }
        set { First.MirrorMargins = value; }
    }

    /// <summary>Where the gutter is (a setting of the document).</summary>
    public GutterPosition GutterPos
    {
        get { return First.GutterPos; }
        set { First.GutterPos = value; }
    }
}
