using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WLock = DocumentFormat.OpenXml.Wordprocessing.Lock;
using Office2010 = DocumentFormat.OpenXml.Office2010.Word;
using Office2013 = DocumentFormat.OpenXml.Office2013.Word;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What a content control holds, as <c>WdContentControlType</c> in Word (with the same values).
/// </summary>
public enum ContentControlType
{
    /// <summary>Text with formatting.</summary>
    RichText = 0,

    /// <summary>Plain text.</summary>
    Text = 1,

    /// <summary>A picture.</summary>
    Picture = 2,

    /// <summary>A list to choose from or type in.</summary>
    ComboBox = 3,

    /// <summary>A list to choose from.</summary>
    DropdownList = 4,

    /// <summary>A building block.</summary>
    BuildingBlockGallery = 5,

    /// <summary>A date.</summary>
    Date = 6,

    /// <summary>A group of other controls.</summary>
    Group = 7,

    /// <summary>A check box.</summary>
    CheckBox = 8,

    /// <summary>A repeating section.</summary>
    RepeatingSection = 9
}

/// <summary>
/// How a content control is drawn, as <c>WdContentControlAppearance</c> in Word (with the same values).
/// </summary>
public enum ContentControlAppearance
{
    /// <summary>With a box around it.</summary>
    BoundingBox = 0,

    /// <summary>With tags at its ends.</summary>
    Tags = 1,

    /// <summary>With no sign of it.</summary>
    Hidden = 2
}

/// <summary>
/// The glossary document Word keeps the placeholders of content controls in, and the rules for the ids of the placeholders.
/// </summary>
internal static class ContentControlParts
{
    internal const string PlaceholderStyle = "PlaceholderText";

    private static string ReadXml(string name)
    {
        using (Stream stream = NewDocument.OpenTemplate(name))
        using (StreamReader reader = new StreamReader(stream))
        {
            string xml = reader.ReadToEnd();
            return xml.Substring(xml.IndexOf("?>", StringComparison.Ordinal) + 2);
        }
    }

    /// <summary>The name of the placeholder that Word gives a kind of control by default.</summary>
    internal static string PlaceholderName(ContentControlType type)
    {
        return type == ContentControlType.ComboBox || type == ContentControlType.DropdownList ? "DefaultPlaceholder_-1854013438"
            : type == ContentControlType.Date ? "DefaultPlaceholder_-1854013437" : "DefaultPlaceholder_-1854013440";
    }

    /// <summary>The text of a placeholder, from the glossary Word wrote.</summary>
    internal static string PlaceholderText(string name)
    {
        DocPart part = new GlossaryDocument(ReadXml("GlossaryDocument.xml")).DocParts.Elements<DocPart>().First(p => p.DocPartProperties.DocPartName.Val.Value == name);
        return part.DocPartBody.InnerText;
    }

    /// <summary>
    /// Takes the placeholders no control uses any more out of the glossary document, and the glossary document itself when none is left, as
    /// Word does when it saves.
    /// </summary>
    internal static void RemoveUnusedPlaceholders(MainDocumentPart main)
    {
        GlossaryDocumentPart glossary = main.GlossaryDocumentPart;
        if (glossary?.GlossaryDocument?.DocParts == null)
        {
            return;
        }
        HashSet<string> used = new HashSet<string>(main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<DocPartReference>()).Select(r => r.Val?.Value));
        foreach (DocPart part in glossary.GlossaryDocument.DocParts.Elements<DocPart>().ToList())
        {
            if (!used.Contains(part.DocPartProperties?.DocPartName?.Val?.Value))
            {
                part.Remove();
            }
        }
        if (!glossary.GlossaryDocument.DocParts.Elements<DocPart>().Any())
        {
            main.DeletePart(glossary);
        }
    }

    /// <summary>
    /// Adds an entry for placeholder text a user wrote, named as Word names it (32 hexadecimal digits), and returns the name. It is made from the
    /// default placeholder entry of Word, with the text changed.
    /// </summary>
    internal static string AddPlaceholder(MainDocumentPart main, string text)
    {
        EnsurePlaceholder(main, "DefaultPlaceholder_-1854013440");
        DocParts parts = main.GlossaryDocumentPart.GlossaryDocument.DocParts;
        DocPart template = parts.Elements<DocPart>().First(p => p.DocPartProperties?.DocPartName?.Val?.Value == "DefaultPlaceholder_-1854013440");
        DocPart copy = (DocPart)template.CloneNode(true);
        string name = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture).ToUpperInvariant();
        copy.DocPartProperties.DocPartName.Val = name;
        copy.DocPartProperties.DocPartId = new DocPartId() { Val = "{" + Guid.NewGuid().ToString("D", CultureInfo.InvariantCulture).ToUpperInvariant() + "}" };
        Text shown = copy.DocPartBody.Descendants<Text>().First();
        foreach (Text extra in copy.DocPartBody.Descendants<Text>().Skip(1).ToList())
        {
            extra.Remove();
        }
        shown.Text = text;
        shown.Space = text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])) ? DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve : null;
        parts.AppendChild(copy);
        return name;
    }

    /// <summary>Makes sure the document has the glossary document and the default placeholder in it.</summary>
    internal static void EnsurePlaceholder(MainDocumentPart main, string name)
    {
        GlossaryDocumentPart glossary = main.GlossaryDocumentPart;
        if (glossary == null)
        {
            glossary = main.AddNewPart<GlossaryDocumentPart>();
            GlossaryDocument document = new GlossaryDocument(ReadXml("GlossaryDocument.xml"));
            document.DocParts.RemoveAllChildren();
            glossary.GlossaryDocument = document;
            glossary.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(ReadXml("GlossaryStyles.xml"));
            glossary.AddNewPart<DocumentSettingsPart>().Settings = new Settings(ReadXml("GlossarySettings.xml"));
            glossary.AddNewPart<WebSettingsPart>().WebSettings = new WebSettings(ReadXml("GlossaryWebSettings.xml"));
            glossary.AddNewPart<FontTablePart>().Fonts = new Fonts(ReadXml("GlossaryFontTable.xml"));
        }

        DocParts parts = glossary.GlossaryDocument.DocParts;
        if (parts.Elements<DocPart>().Any(p => p.DocPartProperties?.DocPartName?.Val?.Value == name))
        {
            return;
        }
        DocPart template = new GlossaryDocument(ReadXml("GlossaryDocument.xml")).DocParts.Elements<DocPart>().First(p => p.DocPartProperties.DocPartName.Val.Value == name);
        DocPart copy = (DocPart)template.CloneNode(true);
        copy.DocPartProperties.DocPartId = new DocPartId() { Val = "{" + Guid.NewGuid().ToString("D", CultureInfo.InvariantCulture).ToUpperInvariant() + "}" };
        parts.AppendChild(copy);
    }
}

/// <summary>
/// A content control, as <c>ContentControl</c> in Word: a <c>w:sdt</c> element.
/// </summary>
public partial class ContentControlWrapper : IEquatable<ContentControlWrapper>
{
    private readonly SdtElement _sdt;

    internal ContentControlWrapper(SdtElement sdt)
    {
        _sdt = sdt;
    }

    internal SdtElement Element
    {
        get { return _sdt; }
    }

    private SdtProperties GetProperties(bool create)
    {
        SdtProperties properties = _sdt.SdtProperties;
        if (properties == null && create)
        {
            properties = new SdtProperties();
            _sdt.PrependChild(properties);
        }
        return properties;
    }

    //The order Word writes the properties in.
    private static int Rank(OpenXmlElement element)
    {
        switch (element.LocalName)
        {
            case "rPr":
                return 0;
            case "alias":
                return 1;
            case "tag":
                return 2;
            case "id":
                return 3;
            case "lock":
                return 4;
            case "placeholder":
                return 5;
            case "temporary":
                return 6;
            case "showingPlcHdr":
                return 7;
            case "dataBinding":
                return 8;
            case "label":
                return 9;
            case "tabIndex":
                return 10;
            case "color":
                return 30;
            case "appearance":
                return 31;
            default:
                return 20;
        }
    }

    internal static void Place(SdtProperties properties, OpenXmlElement element)
    {
        int rank = Rank(element);
        OpenXmlElement before = properties.ChildElements.FirstOrDefault(e => Rank(e) > rank);
        if (before != null)
        {
            before.InsertBeforeSelf(element);
        }
        else
        {
            properties.AppendChild(element);
        }
    }

    private void Set(string localName, OpenXmlElement replacement)
    {
        SdtProperties properties = GetProperties(replacement != null);
        if (properties == null)
        {
            return;
        }
        foreach (OpenXmlElement existing in properties.ChildElements.Where(e => e.LocalName == localName).ToList())
        {
            existing.Remove();
        }
        if (replacement != null)
        {
            Place(properties, replacement);
        }
    }

    private T Find<T>() where T : OpenXmlElement
    {
        return GetProperties(false)?.GetFirstChild<T>();
    }

    /// <summary>What the control holds, as <c>ContentControl.Type</c> in Word.</summary>
    public ContentControlType Type
    {
        get
        {
            SdtProperties properties = GetProperties(false);
            if (properties == null)
            {
                return ContentControlType.RichText;
            }
            foreach (OpenXmlElement element in properties.ChildElements)
            {
                if (element is SdtContentText)
                {
                    return ContentControlType.Text;
                }
                if (element is SdtContentComboBox)
                {
                    return ContentControlType.ComboBox;
                }
                if (element is SdtContentDropDownList)
                {
                    return ContentControlType.DropdownList;
                }
                if (element is SdtContentDate)
                {
                    return ContentControlType.Date;
                }
                if (element is SdtContentPicture)
                {
                    return ContentControlType.Picture;
                }
                if (element is Office2010.SdtContentCheckBox)
                {
                    return ContentControlType.CheckBox;
                }
                if (element is SdtContentDocPartObject)
                {
                    return ContentControlType.BuildingBlockGallery;
                }
                if (element is SdtContentGroup)
                {
                    return ContentControlType.Group;
                }
                if (element.LocalName == "repeatingSection")
                {
                    return ContentControlType.RepeatingSection;
                }
            }
            return ContentControlType.RichText;
        }
    }

    /// <summary>The id of the control, as <c>ContentControl.ID</c> in Word: a number, written as text.</summary>
    public string ID
    {
        get
        {
            int? id = Find<SdtId>()?.Val?.Value;
            return id == null ? string.Empty : unchecked((uint)id.Value).ToString(CultureInfo.InvariantCulture);
        }
    }

    /// <summary>The title shown with the control, as <c>ContentControl.Title</c> in Word.</summary>
    public string Title
    {
        get { return Find<SdtAlias>()?.Val?.Value ?? string.Empty; }
        set { Set("alias", string.IsNullOrEmpty(value) ? null : new SdtAlias() { Val = value }); }
    }

    /// <summary>A tag that identifies the control, as <c>ContentControl.Tag</c> in Word.</summary>
    public string Tag
    {
        get { return Find<Tag>()?.Val?.Value ?? string.Empty; }
        set { Set("tag", string.IsNullOrEmpty(value) ? null : new Tag() { Val = value }); }
    }

    private LockingValues? GetLock()
    {
        return Find<WLock>()?.Val?.Value;
    }

    private void SetLock(bool control, bool contents)
    {
        Set("lock", control && contents ? new WLock() { Val = LockingValues.SdtContentLocked }
            : control ? new WLock() { Val = LockingValues.SdtLocked }
            : contents ? new WLock() { Val = LockingValues.ContentLocked } : null);
    }

    /// <summary>True if the control cannot be deleted, as <c>ContentControl.LockContentControl</c> in Word.</summary>
    public bool LockContentControl
    {
        get { return GetLock() == LockingValues.SdtLocked || GetLock() == LockingValues.SdtContentLocked; }
        set { SetLock(value, LockContents); }
    }

    /// <summary>True if the contents of the control cannot be edited, as <c>ContentControl.LockContents</c> in Word.</summary>
    public bool LockContents
    {
        get { return GetLock() == LockingValues.ContentLocked || GetLock() == LockingValues.SdtContentLocked; }
        set { SetLock(LockContentControl, value); }
    }

    /// <summary>How the control is drawn, as <c>ContentControl.Appearance</c> in Word.</summary>
    public ContentControlAppearance Appearance
    {
        get
        {
            Office2013.SdtAppearance? value = Find<Office2013.Appearance>()?.Val?.Value;
            return value == Office2013.SdtAppearance.Hidden ? ContentControlAppearance.Hidden
                : value == Office2013.SdtAppearance.Tags ? ContentControlAppearance.Tags : ContentControlAppearance.BoundingBox;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Set("appearance", value == ContentControlAppearance.BoundingBox ? null : new Office2013.Appearance()
            {
                Val = value == ContentControlAppearance.Hidden ? Office2013.SdtAppearance.Hidden : Office2013.SdtAppearance.Tags
            });
        }
    }

    /// <summary>True if the control shows its placeholder text, as <c>ContentControl.ShowingPlaceholderText</c> in Word.</summary>
    public bool ShowingPlaceholderText
    {
        get { return Find<ShowingPlaceholder>() != null; }
    }

    /// <summary>
    /// The contents of the control, as <c>ContentControl.Range</c> in Word: the text between its ends.
    /// </summary>
    public RangeWrapper Range
    {
        get
        {
            if (_sdt is SdtRun run)
            {
                StoryMap map = StoryMap.Build(StoryMap.FindRoot(run));
                Atom start = map.Atoms.First(a => a.Kind == AtomKind.FieldMark && ReferenceEquals(a.Element, run));
                SdtContentRun content = run.SdtContentRun;
                Atom end = content == null ? start : map.Atoms.Skip(map.Atoms.IndexOf(start)).First(a => a.Kind == AtomKind.FieldMark && ReferenceEquals(a.Element, content));
                return RangeWrapper.FromMap(map, start.End, content == null ? start.End : end.Start);
            }
            List<Paragraph> paragraphs = _sdt.Descendants<Paragraph>().ToList();
            return paragraphs.Count == 0 ? new RangeWrapper(_sdt) : new RangeWrapper(paragraphs[0], paragraphs[paragraphs.Count - 1]);
        }
    }

    /// <summary>The control this one is inside, as <c>ContentControl.ParentContentControl</c> in Word; null if there is none.</summary>
    public ContentControlWrapper ParentContentControl
    {
        get
        {
            SdtElement parent = _sdt.Ancestors<SdtElement>().FirstOrDefault();
            return parent == null ? null : new ContentControlWrapper(parent);
        }
    }

    /// <summary>
    /// Takes the control away, as <c>ContentControl.Delete</c> in Word. The contents stay in the text unless <paramref name="deleteContents"/>
    /// is true.
    /// </summary>
    public void Delete(bool deleteContents = false)
    {
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(_sdt);
        if (deleteContents)
        {
            if (_sdt is SdtRun)
            {
                RangeWrapper contents = Range;
                if (contents.Start != contents.End)
                {
                    contents.Delete();
                }
            }
            _sdt.Remove();
            ContentControlParts.RemoveUnusedPlaceholders(main);
            return;
        }

        OpenXmlElement content = _sdt.ChildElements.FirstOrDefault(e => e.LocalName == "sdtContent");
        if (content != null)
        {
            foreach (OpenXmlElement child in content.ChildElements.ToList())
            {
                child.Remove();
                _sdt.InsertBeforeSelf(child);
            }
        }
        _sdt.Remove();
        ContentControlParts.RemoveUnusedPlaceholders(main);
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as ContentControlWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(ContentControlWrapper other)
    {
        return !(other is null) && ReferenceEquals(_sdt, other._sdt);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _sdt.GetHashCode();
    }

    public static bool operator ==(ContentControlWrapper left, ContentControlWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(ContentControlWrapper left, ContentControlWrapper right)
    {
        return !(left == right);
    }
}
