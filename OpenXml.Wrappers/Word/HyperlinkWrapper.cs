using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What a hyperlink is attached to, as <c>MsoHyperlinkType</c> in Office (with the same values).
/// </summary>
public enum HyperlinkType
{
    /// <summary>Text.</summary>
    Range = 0,

    /// <summary>A shape.</summary>
    Shape = 1,

    /// <summary>A picture in the text.</summary>
    InlineShape = 2
}

/// <summary>
/// A hyperlink, as <c>Hyperlink</c> in Word: a <c>w:hyperlink</c> element, or a HYPERLINK field, which Word counts as a hyperlink too. The
/// members of a hyperlink made of a field can be read, and it can be deleted, but not changed.
/// </summary>
public class HyperlinkWrapper : IEquatable<HyperlinkWrapper>
{
    private const string StyleId = "Hyperlink";
    private readonly Hyperlink _hyperlink;
    private readonly FieldWrapper _field;
    private readonly DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties _shape;

    internal HyperlinkWrapper(Hyperlink hyperlink)
    {
        _hyperlink = hyperlink;
    }

    internal HyperlinkWrapper(FieldWrapper field)
    {
        _field = field;
    }

    //The link of a shape or picture: a click on it, kept in the properties of its drawing.
    internal HyperlinkWrapper(DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties shape)
    {
        _shape = shape;
    }

    private OpenXmlPart Part
    {
        get { return ((_hyperlink ?? (OpenXmlElement)_shape).GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart; }
    }

    private void ThrowIfShape()
    {
        if (_shape != null)
        {
            throw new NotSupportedException("The hyperlink of a shape has no text, frame or range.");
        }
    }

    //The clicks of a shape: one in the properties of the drawing and, for a picture, one in the picture's own properties.
    private IEnumerable<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick> Clicks
    {
        get
        {
            yield return _shape.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>();
            DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties picture = _shape.Parent?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>().FirstOrDefault();
            if (picture != null)
            {
                yield return picture.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>();
            }
        }
    }

    //The address and the place of the link of a shape are kept together in the target, as address#place.
    private void SetShapeTarget(string address, string place)
    {
        OpenXmlPart part = Part;
        HyperlinkRelationship existing = GetRelationship();
        string tip = _shape.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>()?.Tooltip?.Value;
        if (existing != null)
        {
            part.DeleteReferenceRelationship(existing);
        }
        foreach (DocumentFormat.OpenXml.Drawing.HyperlinkOnClick old in Clicks.Where(c => c != null).ToList())
        {
            old.Remove();
        }
        if (string.IsNullOrEmpty(address) && string.IsNullOrEmpty(place))
        {
            return;
        }
        if (string.IsNullOrEmpty(address))
        {
            throw new NotSupportedException("A shape can go to a place in the document only through an address.");
        }
        string target = string.IsNullOrEmpty(place) ? address : address + "#" + place;
        string id = part.AddHyperlinkRelationship(new Uri(target, UriKind.RelativeOrAbsolute), true).Id;
        DocumentFormat.OpenXml.Drawing.HyperlinkOnClick click = new DocumentFormat.OpenXml.Drawing.HyperlinkOnClick() { Id = id };
        if (!string.IsNullOrEmpty(tip))
        {
            click.Tooltip = tip;
        }
        _shape.HyperlinkOnClick = click;
        DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties picture = _shape.Parent?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>().FirstOrDefault();
        if (picture != null)
        {
            picture.HyperlinkOnClick = (DocumentFormat.OpenXml.Drawing.HyperlinkOnClick)_shape.HyperlinkOnClick.CloneNode(true);
            //Word gives the picture the name of the drawing once it is a link.
            picture.Name = _shape.Name;
        }
    }

    private string ShapeTarget
    {
        get { return GetRelationship()?.Uri.OriginalString ?? string.Empty; }
    }

    private void ThrowIfField()
    {
        if (_field != null)
        {
            throw new NotSupportedException("A hyperlink made of a field cannot be changed.");
        }
    }

    private HyperlinkRelationship GetRelationship()
    {
        string id = _shape != null ? _shape.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>()?.Id?.Value : _hyperlink.Id?.Value;
        return string.IsNullOrEmpty(id) ? null : Part?.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
    }

    //The part of the code of a field after a switch, such as the place after \l.
    private string FieldSwitch(string name)
    {
        string code = _field.CodeText;
        int at = code.IndexOf(name + " \"", StringComparison.Ordinal);
        if (at < 0)
        {
            return string.Empty;
        }
        int start = at + name.Length + 2;
        int end = code.IndexOf('"', start);
        return end < 0 ? string.Empty : code.Substring(start, end - start);
    }

    /// <summary>
    /// The address the hyperlink goes to, as <c>Hyperlink.Address</c> in Word; empty for a hyperlink that goes to a place in the document. For
    /// a mail address with a subject the subject is in plain text, as Word gives it.
    /// </summary>
    public string Address
    {
        get
        {
            if (_field != null)
            {
                string code = _field.CodeText.Trim();
                int first = code.IndexOf('"');
                int end = first < 0 ? -1 : code.IndexOf('"', first + 1);
                return first >= 0 && end > first && !code.StartsWith("\\l", StringComparison.Ordinal) ? code.Substring(first + 1, end - first - 1) : string.Empty;
            }

            string address = GetRelationship()?.Uri.OriginalString ?? string.Empty;
            if (_shape != null)
            {
                int hash = address.IndexOf('#');
                return hash < 0 ? address : address.Substring(0, hash);
            }
            int query = address.IndexOf("?subject=", StringComparison.Ordinal);
            if (query < 0)
            {
                return address;
            }
            string subject = address.Substring(query + "?subject=".Length);
            string unescaped = Uri.UnescapeDataString(subject);
            return string.Format(CultureInfo.InvariantCulture, "{0}?subject={1}", address.Substring(0, query), unescaped);
        }
        set
        {
            ThrowIfField();
            if (_shape != null)
            {
                SetShapeTarget(value, SubAddress);
                return;
            }
            SetAddress(value);
        }
    }

    private void SetAddress(string address)
    {
        OpenXmlPart part = Part;
        HyperlinkRelationship existing = GetRelationship();
        string id = existing?.Id;
        if (existing != null)
        {
            part.DeleteReferenceRelationship(existing);
        }
        if (string.IsNullOrEmpty(address))
        {
            _hyperlink.Id = null;
            return;
        }
        //The new target keeps the id, so that nothing else in the part needs to change.
        HyperlinkRelationship created = id == null ? part.AddHyperlinkRelationship(new Uri(address, UriKind.RelativeOrAbsolute), true)
            : part.AddHyperlinkRelationship(new Uri(address, UriKind.RelativeOrAbsolute), true, id);
        _hyperlink.Id = created.Id;
    }

    /// <summary>The place in the document or the bookmark the hyperlink goes to, as <c>Hyperlink.SubAddress</c> in Word.</summary>
    public string SubAddress
    {
        get
        {
            if (_shape != null)
            {
                string target = ShapeTarget;
                int hash = target.IndexOf('#');
                return hash < 0 ? string.Empty : target.Substring(hash + 1);
            }
            return _field != null ? FieldSwitch("\\l") : _hyperlink.Anchor?.Value ?? string.Empty;
        }
        set
        {
            ThrowIfField();
            if (_shape != null)
            {
                SetShapeTarget(Address, value);
                return;
            }
            _hyperlink.Anchor = string.IsNullOrEmpty(value) ? null : value;
        }
    }

    /// <summary>The text shown when the pointer rests on the hyperlink, as <c>Hyperlink.ScreenTip</c> in Word.</summary>
    public string ScreenTip
    {
        get
        {
            if (_shape != null)
            {
                return _shape.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>()?.Tooltip?.Value ?? string.Empty;
            }
            return _field != null ? FieldSwitch("\\o") : _hyperlink.Tooltip?.Value ?? string.Empty;
        }
        set
        {
            ThrowIfField();
            if (_shape != null)
            {
                foreach (DocumentFormat.OpenXml.Drawing.HyperlinkOnClick click in Clicks.Where(c => c != null))
                {
                    if (string.IsNullOrEmpty(value))
                    {
                        click.Tooltip = null;
                    }
                    else
                    {
                        click.Tooltip = value;
                    }
                }
                return;
            }
            _hyperlink.Tooltip = string.IsNullOrEmpty(value) ? null : value;
        }
    }

    /// <summary>The frame or window the target opens in, as <c>Hyperlink.Target</c> in Word.</summary>
    public string Target
    {
        get { return _shape != null ? string.Empty : _field != null ? FieldSwitch("\\t") : _hyperlink.TargetFrame?.Value ?? string.Empty; }
        set
        {
            ThrowIfField();
            ThrowIfShape();
            _hyperlink.TargetFrame = string.IsNullOrEmpty(value) ? null : value;
        }
    }

    /// <summary>
    /// The subject of the message a mail hyperlink starts, as <c>Hyperlink.EmailSubject</c> in Word; it is part of <see cref="Address"/>.
    /// </summary>
    public string EmailSubject
    {
        get
        {
            string address = Address;
            int query = address.IndexOf("?subject=", StringComparison.Ordinal);
            return query < 0 ? string.Empty : address.Substring(query + "?subject=".Length);
        }
        set
        {
            ThrowIfField();
            ThrowIfShape();
            string address = GetRelationship()?.Uri.OriginalString ?? string.Empty;
            int query = address.IndexOf('?');
            if (query >= 0)
            {
                address = address.Substring(0, query);
            }
            SetAddress(string.IsNullOrEmpty(value) ? address : address + "?subject=" + Uri.EscapeDataString(value));
        }
    }

    /// <summary>The address and place the hyperlink goes to, as <c>Hyperlink.Name</c> in Word: the address, then <c> - </c> and the place.</summary>
    public string Name
    {
        get
        {
            string address = Address;
            string place = SubAddress;
            return address.Length == 0 ? place : place.Length == 0 ? address : address + " - " + place;
        }
    }

    /// <summary>What the hyperlink is attached to, as <c>Hyperlink.Type</c> in Word; always <see cref="HyperlinkType.Range"/> here.</summary>
    public HyperlinkType Type
    {
        get { return _shape == null ? HyperlinkType.Range : _shape.Parent is DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline ? HyperlinkType.InlineShape : HyperlinkType.Shape; }
    }

    /// <summary>
    /// The text the hyperlink shows, as <c>Hyperlink.TextToDisplay</c> in Word. Setting it replaces the text with new text in the Hyperlink style.
    /// </summary>
    public string TextToDisplay
    {
        get
        {
            if (_shape != null)
            {
                return string.Empty;
            }
            if (_field != null)
            {
                return _field.Result.Text;
            }
            RangeWrapper content = GetContentRange();
            return content.Start == content.End ? string.Empty : content.Text;
        }
        set
        {
            ThrowIfField();
            ThrowIfShape();
            string text = value ?? string.Empty;
            if (!_hyperlink.Elements<Run>().Any())
            {
                _hyperlink.AppendChild(NewRun(text));
                return;
            }
            //The new text goes in first so that the hyperlink is never empty while the old text is deleted.
            RangeWrapper old = GetContentRange();
            _hyperlink.AppendChild(NewRun(text));
            old.Delete();
        }
    }

    /// <summary>The range of the hyperlink, as <c>Hyperlink.Range</c> in Word: its text, with the characters that Word counts for it.</summary>
    public RangeWrapper Range
    {
        get
        {
            ThrowIfShape();
            if (_field != null)
            {
                return _field.Range;
            }
            GetSpan(out StoryMap map, out int start, out int end, false);
            return RangeWrapper.FromMap(map, start, end);
        }
    }

    /// <summary>A range over the paragraph or other element the hyperlink is in, as <c>Hyperlink.Parent</c> in Word for text.</summary>
    public RangeWrapper Parent
    {
        get
        {
            ThrowIfShape();
            return new RangeWrapper(_field != null ? _field.Range.AnchorElement.Parent : _hyperlink.Parent);
        }
    }

    //The span of the hyperlink: its code, its text and its end, or only the text.
    internal void GetSpan(out StoryMap map, out int start, out int end, bool textOnly)
    {
        map = StoryMap.Build(StoryMap.FindRoot(_hyperlink));
        int code = map.Atoms.FindIndex(a => a.Kind == AtomKind.FieldCode && ReferenceEquals(a.Element, _hyperlink));
        Atom first = map.Atoms[code];
        OpenXmlElement last = _hyperlink.LastChild;
        Atom close = last == null ? first : map.Atoms.Skip(code).First(a => a.Kind == AtomKind.FieldMark && ReferenceEquals(a.Element, last));
        start = textOnly ? first.End : first.Start;
        end = textOnly ? (last == null ? first.End : close.Start) : close.End;
    }

    private RangeWrapper GetContentRange()
    {
        GetSpan(out StoryMap map, out int start, out int end, true);
        return RangeWrapper.FromMap(map, start, end);
    }

    internal static Run NewRun(string text)
    {
        Run run = new Run(new RunProperties(new RunStyle() { Val = StyleId }));
        if (text.Length > 0)
        {
            run.AppendChild(OpenXmlExtensions.NewText(text));
        }
        return run;
    }

    //Gives a run the Hyperlink character style, first among its properties.
    internal static void ApplyStyle(Run run)
    {
        RunProperties properties = run.GetOrAddRunProperties();
        properties.RemoveAllChildren<RunStyle>();
        properties.PrependChild(new RunStyle() { Val = StyleId });
    }

    /// <summary>
    /// Takes the hyperlink away and leaves its text, as <c>Hyperlink.Delete</c> in Word: the text is no longer in the Hyperlink style, and the
    /// address is dropped from the document.
    /// </summary>
    public void Delete()
    {
        if (_field != null)
        {
            _field.Unlink();
            return;
        }
        if (_shape != null)
        {
            SetShapeTarget(null, null);
            return;
        }

        foreach (Run run in _hyperlink.Descendants<Run>().ToList())
        {
            RunProperties properties = run.RunProperties;
            RunStyle style = properties?.RunStyle;
            if (style != null && style.Val?.Value == StyleId)
            {
                style.Remove();
                if (!properties.HasChildren)
                {
                    properties.Remove();
                }
            }
        }
        HyperlinkRelationship relationship = GetRelationship();
        OpenXmlPart part = Part;
        foreach (OpenXmlElement child in _hyperlink.ChildElements.ToList())
        {
            child.Remove();
            _hyperlink.InsertBeforeSelf(child);
        }
        _hyperlink.Remove();
        if (relationship != null && part != null && !(part.RootElement?.Descendants<Hyperlink>().Any(h => h.Id?.Value == relationship.Id) ?? false))
        {
            part.DeleteReferenceRelationship(relationship);
        }
    }

    #region Equality

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as HyperlinkWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(HyperlinkWrapper other)
    {
        return !(other is null) && ReferenceEquals(Key, other.Key);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return Key.GetHashCode();
    }

    private object Key
    {
        get { return _hyperlink != null ? (object)_hyperlink : _shape != null ? _shape : _field.Key; }
    }

    public static bool operator ==(HyperlinkWrapper left, HyperlinkWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(HyperlinkWrapper left, HyperlinkWrapper right)
    {
        return !(left == right);
    }

    #endregion
}

/// <summary>
/// The hyperlinks of a range or a document, as <c>Hyperlinks</c> in Word: <c>w:hyperlink</c> elements and HYPERLINK fields, in the order they
/// are in the text.
/// </summary>
public partial class HyperlinkWrapperCollection : IEnumerable<HyperlinkWrapper>
{
    private readonly RangeWrapper _range;

    internal HyperlinkWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range the hyperlinks are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<HyperlinkWrapper> GetItems()
    {
        OpenXmlElement root = StoryMap.FindRoot(_range.AnchorElement);
        StoryMap map = StoryMap.Build(root);
        _range.GetStoryPositions(out int rangeStart, out int rangeEnd);
        List<HyperlinkWrapper> items = new List<HyperlinkWrapper>();
        foreach (Atom atom in map.Atoms)
        {
            HyperlinkWrapper item = null;
            int start = atom.Start;
            int end = atom.End;
            if (atom.Kind == AtomKind.FieldCode && atom.Element is Hyperlink link)
            {
                item = new HyperlinkWrapper(link);
                item.GetSpan(out _, out start, out end, false);
            }
            else if (atom.Kind == AtomKind.FieldCode && atom.Element is SimpleField simple && simple.Instruction?.Value?.TrimStart().StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase) == true)
            {
                FieldWrapper field = new FieldWrapper(simple);
                item = new HyperlinkWrapper(field);
                field.Range.GetStoryPositions(out start, out end);
            }
            else if (atom.Kind == AtomKind.FieldMark && atom.Element is FieldChar begin && begin.FieldCharType?.Value == FieldCharValues.Begin)
            {
                FieldWrapper field = new FieldWrapper(begin);
                if (field.Type == FieldType.Hyperlink)
                {
                    item = new HyperlinkWrapper(field);
                    field.Range.GetStoryPositions(out start, out end);
                }
            }

            bool inRange = rangeStart == rangeEnd ? start <= rangeStart && rangeStart <= end : start < rangeEnd && end > rangeStart;
            if (item != null && inRange)
            {
                items.Add(item);
            }
        }
        return items;
    }

    /// <summary>The number of hyperlinks.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The hyperlink at the index (0-based), as <c>Hyperlinks.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no hyperlink at the index.</exception>
    public HyperlinkWrapper this[int index]
    {
        get
        {
            List<HyperlinkWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>
    /// Makes the text of a range a hyperlink, as <c>Hyperlinks.Add</c> in Word, and returns it. A range with no text gets
    /// <paramref name="textToDisplay"/>, or the address, as its text. The text is replaced by <paramref name="textToDisplay"/> when it is given.
    /// </summary>
    /// <param name="anchor">The range to make a hyperlink of. It must be in one paragraph.</param>
    /// <param name="address">The address to go to; empty for a place in the document.</param>
    /// <param name="subAddress">The place in the document (a bookmark) or in the address to go to.</param>
    /// <param name="screenTip">The text to show when the pointer rests on the hyperlink.</param>
    /// <param name="textToDisplay">The text of the hyperlink.</param>
    /// <param name="target">The frame or window to open the address in.</param>
    /// <exception cref="ArgumentException">There is neither an address nor a place.</exception>
    /// <exception cref="NotSupportedException">The range is not in one paragraph, or is inside another hyperlink.</exception>
    public HyperlinkWrapper Add(RangeWrapper anchor, string address = null, string subAddress = null, string screenTip = null, string textToDisplay = null, string target = null)
    {
        Guard.ThrowIfNull(anchor);
        if (string.IsNullOrEmpty(address) && string.IsNullOrEmpty(subAddress))
        {
            throw new ArgumentException("A hyperlink needs an address or a place to go to.", nameof(address));
        }

        RangeWrapper work = anchor.Duplicate();
        anchor.GetStoryPositions(out int start, out int end);
        if (start == end)
        {
            work.InsertAfter(textToDisplay ?? (string.IsNullOrEmpty(address) ? subAddress : address));
        }
        else if (textToDisplay != null)
        {
            work.Text = textToDisplay;
        }

        RangeWrapper elements = work.ElementRange();
        if (!(elements.StartElement is Run first) || !(elements.EndElement is Run last) || !(first.Parent is Paragraph paragraph) || !ReferenceEquals(paragraph, last.Parent))
        {
            throw new NotSupportedException("A hyperlink can be made only of the runs of one paragraph.");
        }

        MainDocumentPart main = StyleHelper.GetMainDocumentPart(paragraph);
        BuiltInStyles.GetOrCreate(main, "Hyperlink");
        Hyperlink link = new Hyperlink();
        //Word writes 1 where the SDK would write true.
        link.SetAttribute(new OpenXmlAttribute("w", "history", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "1"));
        OpenXmlPart part = (paragraph.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart;
        if (!string.IsNullOrEmpty(address))
        {
            link.Id = part.AddHyperlinkRelationship(new Uri(address, UriKind.RelativeOrAbsolute), true).Id;
        }
        if (!string.IsNullOrEmpty(subAddress))
        {
            link.Anchor = subAddress;
        }
        if (!string.IsNullOrEmpty(target))
        {
            link.TargetFrame = target;
        }
        if (!string.IsNullOrEmpty(screenTip))
        {
            link.Tooltip = screenTip;
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
        first.InsertBeforeSelf(link);
        foreach (OpenXmlElement element in moving)
        {
            element.Remove();
            link.AppendChild(element);
            if (element is Run run)
            {
                HyperlinkWrapper.ApplyStyle(run);
            }
        }
        return new HyperlinkWrapper(link);
    }

    /// <inheritdoc/>
    public IEnumerator<HyperlinkWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
