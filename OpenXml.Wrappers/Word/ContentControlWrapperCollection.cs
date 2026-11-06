using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WCalendar = DocumentFormat.OpenXml.Wordprocessing.Calendar;
using Office2010 = DocumentFormat.OpenXml.Office2010.Word;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The content controls of a range or a document, as <c>ContentControls</c> in Word.
/// </summary>
public class ContentControlWrapperCollection : IEnumerable<ContentControlWrapper>
{
    private const string W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private readonly RangeWrapper _range;
    private readonly Func<ContentControlWrapper, bool> _filter;

    internal ContentControlWrapperCollection(RangeWrapper range, Func<ContentControlWrapper, bool> filter = null)
    {
        _range = range;
        _filter = filter;
    }

    /// <summary>The range the controls are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<ContentControlWrapper> GetItems()
    {
        OpenXmlElement root = StoryMap.FindRoot(_range.AnchorElement);
        StoryMap map = StoryMap.Build(root);
        _range.GetStoryPositions(out int rangeStart, out int rangeEnd);
        bool whole = rangeStart <= 0 && rangeEnd >= map.Length;
        List<ContentControlWrapper> items = new List<ContentControlWrapper>();
        foreach (SdtElement sdt in root.Descendants<SdtElement>())
        {
            bool inRange = whole;
            if (!whole && sdt is SdtRun run)
            {
                Atom start = map.Atoms.FirstOrDefault(a => a.Kind == AtomKind.FieldMark && ReferenceEquals(a.Element, run));
                SdtContentRun content = run.SdtContentRun;
                Atom end = content == null || start == null ? start : map.Atoms.Skip(map.Atoms.IndexOf(start)).FirstOrDefault(a => a.Kind == AtomKind.FieldMark && ReferenceEquals(a.Element, content));
                if (start != null)
                {
                    int from = start.Start;
                    int to = end == null ? start.End : end.End;
                    inRange = rangeStart == rangeEnd ? from <= rangeStart && rangeStart <= to : from < rangeEnd && to > rangeStart;
                }
            }
            if (inRange)
            {
                ContentControlWrapper wrapper = new ContentControlWrapper(sdt);
                if (_filter == null || _filter(wrapper))
                {
                    items.Add(wrapper);
                }
            }
        }
        return items;
    }

    /// <summary>The number of controls.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The control at the index (0-based), as <c>ContentControls.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no control at the index.</exception>
    public ContentControlWrapper this[int index]
    {
        get
        {
            List<ContentControlWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>The innermost control whose contents hold the range.</summary>
    internal static ContentControlWrapper FindParent(RangeWrapper range)
    {
        OpenXmlElement root = StoryMap.FindRoot(range.AnchorElement);
        SdtElement byElement = range.IsSpan ? null : range.AnchorElement.Ancestors<SdtElement>().FirstOrDefault();
        if (byElement != null)
        {
            return new ContentControlWrapper(byElement);
        }
        StoryMap map = StoryMap.Build(root);
        range.GetStoryPositions(out int rangeStart, out int rangeEnd);
        ContentControlWrapper best = null;
        int bestLength = int.MaxValue;
        foreach (SdtRun run in root.Descendants<SdtRun>())
        {
            Atom start = map.Atoms.FirstOrDefault(a => a.Kind == AtomKind.FieldMark && ReferenceEquals(a.Element, run));
            SdtContentRun content = run.SdtContentRun;
            if (start == null || content == null)
            {
                continue;
            }
            Atom end = map.Atoms.Skip(map.Atoms.IndexOf(start)).FirstOrDefault(a => a.Kind == AtomKind.FieldMark && ReferenceEquals(a.Element, content));
            if (end != null && rangeStart >= start.End && rangeEnd <= end.Start && end.Start - start.End < bestLength)
            {
                best = new ContentControlWrapper(run);
                bestLength = end.Start - start.End;
            }
        }
        return best;
    }

    private static Run PlaceholderRun(string text)
    {
        return new Run(new RunProperties(new RunStyle() { Val = ContentControlParts.PlaceholderStyle }), OpenXmlExtensions.NewText(text));
    }

    private static OpenXmlElement TypeElement(ContentControlType type)
    {
        switch (type)
        {
            case ContentControlType.Text:
                return new SdtContentText();
            case ContentControlType.ComboBox:
                return new SdtContentComboBox();
            case ContentControlType.DropdownList:
                return new SdtContentDropDownList();
            case ContentControlType.Date:
                return new SdtContentDate(
                    new DateFormat() { Val = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern },
                    new LanguageId() { Val = CultureInfo.CurrentCulture.Name },
                    new SdtDateMappingType() { Val = DateFormatValues.DateTime },
                    new WCalendar() { Val = CalendarValues.Gregorian });
            case ContentControlType.CheckBox:
                Office2010.SdtContentCheckBox box = new Office2010.SdtContentCheckBox();
                box.AppendChild(WithAttributes(new Office2010.Checked(), ("val", "0")));
                box.AppendChild(WithAttributes(new Office2010.CheckedState(), ("font", "MS Gothic"), ("val", "2612")));
                box.AppendChild(WithAttributes(new Office2010.UncheckedState(), ("font", "MS Gothic"), ("val", "2610")));
                return box;
            default:
                return null;
        }
    }

    private static OpenXmlElement WithAttributes(OpenXmlElement element, params (string Name, string Value)[] attributes)
    {
        foreach ((string name, string value) in attributes)
        {
            element.SetAttribute(new OpenXmlAttribute("w14", name, W14, value));
        }
        return element;
    }

    /// <summary>
    /// Makes a content control of a range, as <c>ContentControls.Add</c> in Word. A control around text holds the text; one at a point holds
    /// its placeholder text (or, for a check box, an empty box).
    /// </summary>
    /// <param name="type">The kind of control: rich text, text, combo box, drop-down list, date or check box.</param>
    /// <param name="range">The text to put the control around, or the place for it.</param>
    /// <exception cref="NotSupportedException">The kind of control is not one of those, the range is not in one paragraph, or a check box is made around text.</exception>
    public ContentControlWrapper Add(ContentControlType type, RangeWrapper range)
    {
        Guard.ThrowIfNull(range);
        if (type != ContentControlType.RichText && type != ContentControlType.Text && type != ContentControlType.ComboBox
            && type != ContentControlType.DropdownList && type != ContentControlType.Date && type != ContentControlType.CheckBox)
        {
            throw new NotSupportedException("Only rich text, text, combo box, drop-down list, date and check box controls can be made.");
        }

        range.GetStoryPositions(out int start, out int end);
        bool collapsed = start == end;
        if (type == ContentControlType.CheckBox && !collapsed)
        {
            throw new NotSupportedException("A check box is made at a point, not around text.");
        }
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(range.AnchorElement);
        if (main == null)
        {
            throw new NotSupportedException("A content control can be made only in a document.");
        }

        SdtProperties properties = new SdtProperties();
        properties.AppendChild(new SdtId() { Val = ListDefaults.NewRandom() });
        if (type != ContentControlType.CheckBox)
        {
            string name = ContentControlParts.PlaceholderName(type);
            ContentControlParts.EnsurePlaceholder(main, name);
            properties.AppendChild(new SdtPlaceholder(new DocPartReference() { Val = name }));
            if (collapsed)
            {
                properties.AppendChild(new ShowingPlaceholder());
            }
        }
        OpenXmlElement typeElement = TypeElement(type);
        if (typeElement != null)
        {
            properties.AppendChild(typeElement);
        }

        SdtRun sdt = new SdtRun(properties);
        SdtContentRun content = new SdtContentRun();
        sdt.AppendChild(content);

        if (collapsed)
        {
            if (type == ContentControlType.CheckBox)
            {
                content.AppendChild(new Run(
                    new RunProperties(new RunFonts() { Ascii = "MS Gothic", EastAsia = "MS Gothic", HighAnsi = "MS Gothic", Hint = FontTypeHintValues.EastAsia }),
                    OpenXmlExtensions.NewText("☐")));
            }
            else
            {
                BuiltInStyles.GetOrCreate(main, ContentControlParts.PlaceholderStyle);
                content.AppendChild(PlaceholderRun(ContentControlParts.PlaceholderText(ContentControlParts.PlaceholderName(type))));
            }

            Run marker = new Run();
            RangeWrapper at = range.Duplicate();
            at.InsertNoteReference(paragraph => new List<Run> { marker });
            marker.InsertBeforeSelf(sdt);
            marker.Remove();
        }
        else
        {
            RangeWrapper elements = range.Duplicate().ElementRange();
            if (!(elements.StartElement is Run first) || !(elements.EndElement is Run last) || !(first.Parent is Paragraph paragraph) || !ReferenceEquals(paragraph, last.Parent))
            {
                throw new NotSupportedException("A content control can be made only around the runs of one paragraph.");
            }
            List<OpenXmlElement> moving = new List<OpenXmlElement>();
            for (OpenXmlElement element = first; element != null; element = element.NextSibling())
            {
                moving.Add(element);
                if (ReferenceEquals(element, last))
                {
                    break;
                }
            }
            first.InsertBeforeSelf(sdt);
            foreach (OpenXmlElement element in moving)
            {
                element.Remove();
                content.AppendChild(element);
            }
        }
        return new ContentControlWrapper(sdt);
    }

    /// <inheritdoc/>
    public IEnumerator<ContentControlWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
