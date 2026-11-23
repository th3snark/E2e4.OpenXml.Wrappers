using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The page number shown at a place in the main text, estimated without laying the pages out: a page starts after a page break, at a
/// paragraph that starts on a new page, after a section break that starts the next section on a new page, and where Word last broke a page
/// (<c>w:lastRenderedPageBreak</c>, which a document saved by Word has). The number follows each section's numbering (its starting number and
/// format). A document Word last saved gives Word's numbers; otherwise they are as far as the breaks in the text tell.
/// </summary>
internal sealed class PageEstimator
{
    private readonly Body _body;
    private readonly Dictionary<OpenXmlElement, int> _pages = new Dictionary<OpenXmlElement, int>();
    private readonly Dictionary<OpenXmlElement, int> _sections = new Dictionary<OpenXmlElement, int>();
    private readonly List<int> _firstPages = new List<int>();
    private readonly List<SectionProperties> _properties;
    private StoryMap _map;

    public PageEstimator(Body body)
    {
        _body = body;
        _properties = body.Descendants<Paragraph>().Select(p => p.ParagraphProperties?.SectionProperties).Where(s => s != null).ToList();
        SectionProperties last = body.GetFirstChild<SectionProperties>();
        if (last != null)
        {
            _properties.Add(last);
        }
        Walk();
    }

    private void Walk()
    {
        int page = 1;
        int section = 0;
        bool atPageStart = true;
        bool sectionEnded = false;
        _firstPages.Add(1);
        foreach (OpenXmlElement element in _body.Descendants())
        {
            if (element is Paragraph paragraph && !paragraph.Ancestors<TextBoxContent>().Any())
            {
                if (sectionEnded)
                {
                    //The paragraph before ended a section: the next one starts on a new page unless it is continuous.
                    section++;
                    SectionProperties next = section < _properties.Count ? _properties[section] : null;
                    SectionMarkValues? type = next?.GetFirstChild<SectionType>()?.Val?.Value;
                    if (type != SectionMarkValues.Continuous && type != SectionMarkValues.NextColumn)
                    {
                        page++;
                        atPageStart = true;
                    }
                    _firstPages.Add(page);
                    sectionEnded = false;
                }
                if (paragraph.ParagraphProperties?.PageBreakBefore != null && OnOff(paragraph.ParagraphProperties.PageBreakBefore.Val) && !atPageStart)
                {
                    page++;
                    atPageStart = true;
                }
                sectionEnded = paragraph.ParagraphProperties?.SectionProperties != null;
            }
            if (element is LastRenderedPageBreak)
            {
                if (!atPageStart)
                {
                    page++;
                }
                atPageStart = true;
            }
            _pages[element] = page;
            _sections[element] = section;

            if (element is Break pageBreak && pageBreak.Type != null && pageBreak.Type.HasValue && pageBreak.Type.Value == BreakValues.Page)
            {
                page++;
                atPageStart = true;
            }
            else if (element is Text || element is TabChar || element is Drawing || element is Picture || element is FootnoteReference || element is EndnoteReference)
            {
                atPageStart = false;
            }
        }
    }

    private static bool OnOff(OnOffValue value)
    {
        return value == null || !value.HasValue || value.Value;
    }

    /// <summary>The page number shown at the start of the range: in a note or a text box, where its mark or anchor is.</summary>
    public string PageNumberAt(RangeWrapper range)
    {
        OpenXmlElement element;
        if (StoryMap.FindRoot(range.AnchorElement) is Body)
        {
            _map = _map ?? StoryMap.Build(_body);
            int index = _map.IndexAt(range.Start);
            element = index >= 0 ? _map.Atoms[index].Element : _body.Descendants().LastOrDefault();
        }
        else
        {
            element = RangeWrapper.SectionAnchor(range.AnchorElement, 0);
        }
        if (element == null || !_pages.TryGetValue(element, out int page))
        {
            return null;
        }
        int section = _sections[element];
        return Format(Number(page, section), section);
    }

    //The number of a page: from the section's starting number when it has one, otherwise going on from the section before.
    private int Number(int page, int section)
    {
        int number = page;
        int shift = 0;
        for (int s = 0; s <= section && s < _firstPages.Count; s++)
        {
            int? start = StartOf(s);
            if (start.HasValue)
            {
                shift = start.Value - _firstPages[s];
            }
        }
        return number + shift;
    }

    private int? StartOf(int section)
    {
        PageNumberType numbering = section < _properties.Count ? _properties[section].GetFirstChild<PageNumberType>() : null;
        return numbering?.Start?.Value;
    }

    private string Format(int number, int section)
    {
        PageNumberType numbering = section < _properties.Count ? _properties[section].GetFirstChild<PageNumberType>() : null;
        NumberFormatValues? format = numbering?.Format?.Value;
        string name = format == NumberFormatValues.UpperRoman ? "ROMAN"
            : format == NumberFormatValues.LowerRoman ? "roman"
            : format == NumberFormatValues.UpperLetter ? "ALPHABETIC"
            : format == NumberFormatValues.LowerLetter ? "alphabetic"
            : format == NumberFormatValues.NumberInDash ? "ArabicDash"
            : "ARABIC";
        return FieldNumberFormats.Format(number, name);
    }
}
