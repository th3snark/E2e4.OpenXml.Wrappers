using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What kind of file a document is, as <c>WdDocumentType</c> in Word (with the same values; named so as not to clash with the SDK's <c>DocumentType</c>).
/// </summary>
public enum WordDocumentType
{
    /// <summary>A document.</summary>
    Document = 0,

    /// <summary>A template.</summary>
    Template = 1
}

public partial class WordprocessingDocumentWrapper
{
    private bool GetSetting<T>() where T : OnOffType
    {
        T setting = GetSettings(false)?.GetFirstChild<T>();
        return setting != null && (setting.Val == null || !setting.Val.HasValue || setting.Val.Value);
    }

    //Word writes the setting when it is on and leaves it out when it is off.
    private void SetSetting<T>(bool value) where T : OnOffType, new()
    {
        if (value)
        {
            Settings settings = GetSettings(true);
            if (!GetSetting<T>())
            {
                settings.RemoveAllChildren<T>();
                settings.AddChild(new T());
            }
        }
        else
        {
            GetSettings(false)?.RemoveAllChildren<T>();
        }
    }

    /// <summary>
    /// The distance between the default tab stops, in points, as <c>Document.DefaultTabStop</c> in Word: 36 (half an inch) unless the document says
    /// otherwise.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The distance is not more than 0, or is more than 1638 points (the most the SDK can store).</exception>
    public double DefaultTabStop
    {
        get
        {
            long? twips = GetSettings(false)?.GetFirstChild<DefaultTabStop>()?.Val?.Value;
            return twips.HasValue ? twips.Value / 20.0 : 36;
        }
        set
        {
            if (!(value > 0 && value <= 1638))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The distance must be more than 0 and at most 1638 points.");
            }

            Settings settings = GetSettings(true);
            settings.RemoveAllChildren<DefaultTabStop>();
            settings.AddChild(new DefaultTabStop() { Val = (short)Math.Min(Math.Round(value * 20), short.MaxValue) });
        }
    }

    /// <summary>True to hyphenate the document automatically, as <c>Document.AutoHyphenation</c> in Word.</summary>
    public bool AutoHyphenation
    {
        get { return GetSetting<AutoHyphenation>(); }
        set { SetSetting<AutoHyphenation>(value); }
    }

    /// <summary>True (the default) to hyphenate words in capitals, as <c>Document.HyphenateCaps</c> in Word.</summary>
    public bool HyphenateCaps
    {
        get { return !GetSetting<DoNotHyphenateCaps>(); }
        set { SetSetting<DoNotHyphenateCaps>(!value); }
    }

    /// <summary>
    /// The most lines in a row that may end in a hyphen, as <c>Document.ConsecutiveHyphensLimit</c> in Word; 0 for no limit.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The limit is less than 0.</exception>
    public int ConsecutiveHyphensLimit
    {
        get { return GetSettings(false)?.GetFirstChild<ConsecutiveHyphenLimit>()?.Val?.Value ?? 0; }
        set
        {
            Guard.ThrowIfNegative(value);

            Settings settings = GetSettings(value > 0);
            settings?.RemoveAllChildren<ConsecutiveHyphenLimit>();
            if (value > 0)
            {
                settings.AddChild(new ConsecutiveHyphenLimit() { Val = (ushort)Math.Min(value, ushort.MaxValue) });
            }
        }
    }

    /// <summary>True if the styles of the attached template are copied into the document when it is opened, as <c>Document.UpdateStylesOnOpen</c> in Word (<c>w:linkStyles</c>).</summary>
    public bool UpdateStylesOnOpen
    {
        get { return GetSetting<LinkStyles>(); }
        set { SetSetting<LinkStyles>(value); }
    }

    /// <summary>True to save the fonts the document uses in it, as <c>Document.EmbedTrueTypeFonts</c> in Word. The library only keeps the setting.</summary>
    public bool EmbedTrueTypeFonts
    {
        get { return GetSetting<EmbedTrueTypeFonts>(); }
        set { SetSetting<EmbedTrueTypeFonts>(value); }
    }

    /// <summary>True to save only the characters of embedded fonts that the document uses, as <c>Document.SaveSubsetFonts</c> in Word. The library only keeps the setting.</summary>
    public bool SaveSubsetFonts
    {
        get { return GetSetting<SaveSubsetFonts>(); }
        set { SetSetting<SaveSubsetFonts>(value); }
    }

    /// <summary>True (the default) to leave out the system fonts when fonts are embedded, as <c>Document.DoNotEmbedSystemFonts</c> in Word.</summary>
    public bool DoNotEmbedSystemFonts
    {
        get { return !GetSetting<EmbedSystemFonts>(); }
        set { SetSetting<EmbedSystemFonts>(!value); }
    }

    /// <summary>True to take the author's name and similar information out of the document when Word saves it, as <c>Document.RemovePersonalInformation</c> in Word. The library only keeps the setting.</summary>
    public bool RemovePersonalInformation
    {
        get { return GetSetting<RemovePersonalInformation>(); }
        set { SetSetting<RemovePersonalInformation>(value); }
    }

    /// <summary>True to take the dates and times of comments out of the document when Word saves it, as <c>Document.RemoveDateAndTime</c> in Word. The library only keeps the setting.</summary>
    public bool RemoveDateAndTime
    {
        get { return GetSetting<RemoveDateAndTime>(); }
        set { SetSetting<RemoveDateAndTime>(value); }
    }

    /// <summary>Whether the document is a document or a template, as <c>Document.Type</c> in Word.</summary>
    public WordDocumentType Type
    {
        get
        {
            WordprocessingDocumentType type = _wordprocessingDocument.DocumentType;
            return type == WordprocessingDocumentType.Template || type == WordprocessingDocumentType.MacroEnabledTemplate ? WordDocumentType.Template : WordDocumentType.Document;
        }
    }

    /// <summary>The format the document is in (.docx, .docm, .dotx or .dotm), as <c>Document.SaveFormat</c> in Word.</summary>
    public SaveFormat SaveFormat
    {
        get
        {
            WordprocessingDocumentType type = _wordprocessingDocument.DocumentType;
            return type == WordprocessingDocumentType.MacroEnabledDocument ? SaveFormat.XmlDocumentMacroEnabled
                : type == WordprocessingDocumentType.Template ? SaveFormat.XmlTemplate
                : type == WordprocessingDocumentType.MacroEnabledTemplate ? SaveFormat.XmlTemplateMacroEnabled : SaveFormat.XmlDocument;
        }
    }

    /// <summary>The content controls in the main text, as <c>Document.ContentControls</c> in Word.</summary>
    public ContentControlWrapperCollection ContentControls
    {
        get { return new ContentControlWrapperCollection(Content); }
    }

    /// <summary>The floating shapes in the main text, as <c>Document.Shapes</c> in Word.</summary>
    public ShapeWrapperCollection Shapes
    {
        get { return new ShapeWrapperCollection(Content); }
    }

    /// <summary>The tables of contents in the main text, as <c>Document.TablesOfContents</c> in Word.</summary>
    public TableOfContentsWrapperCollection TablesOfContents
    {
        get { return new TableOfContentsWrapperCollection(Content); }
    }

    /// <summary>The tables of figures in the main text, as <c>Document.TablesOfFigures</c> in Word.</summary>
    public TableOfFiguresWrapperCollection TablesOfFigures
    {
        get { return new TableOfFiguresWrapperCollection(Content); }
    }

    /// <summary>The inline shapes (pictures) in the main text, as <c>Document.InlineShapes</c> in Word.</summary>
    public InlineShapeWrapperCollection InlineShapes
    {
        get { return new InlineShapeWrapperCollection(Content); }
    }

    /// <summary>The custom XML parts of the document, as <c>Document.CustomXMLParts</c> in Word.</summary>
    public CustomXMLParts CustomXMLParts
    {
        get { return new CustomXMLParts(WordprocessingDocument.MainDocumentPart); }
    }

    /// <summary>The content controls with a tag, as <c>Document.SelectContentControlsByTag</c> in Word.</summary>
    public ContentControlWrapperCollection SelectContentControlsByTag(string tag)
    {
        return new ContentControlWrapperCollection(Content, c => string.Equals(c.Tag, tag, System.StringComparison.Ordinal));
    }

    /// <summary>The content controls with a title, as <c>Document.SelectContentControlsByTitle</c> in Word.</summary>
    public ContentControlWrapperCollection SelectContentControlsByTitle(string title)
    {
        return new ContentControlWrapperCollection(Content, c => string.Equals(c.Title, title, System.StringComparison.Ordinal));
    }

    /// <summary>The content controls bound to a node of a custom XML part, as <c>Document.SelectLinkedControls</c> in Word.</summary>
    /// <exception cref="System.ArgumentNullException">The node is null.</exception>
    public ContentControlWrapperCollection SelectLinkedControls(CustomXMLNode node)
    {
        Guard.ThrowIfNull(node);
        return new ContentControlWrapperCollection(Content, c => c.XMLMapping.CustomXMLNode == node);
    }

    /// <summary>The content controls that are not bound to a custom XML part, as <c>Document.SelectUnlinkedControls</c> in Word.</summary>
    /// <param name="stream">A part; the controls bound to other parts or to none are listed.</param>
    public ContentControlWrapperCollection SelectUnlinkedControls(CustomXMLPart stream = null)
    {
        return new ContentControlWrapperCollection(Content, c => !c.XMLMapping.IsMapped || (stream != null && c.XMLMapping.CustomXMLPart != stream));
    }

    /// <summary>The footnotes of the main text, as <c>Document.Footnotes</c> in Word.</summary>
    public FootnoteWrapperCollection Footnotes
    {
        get { return new FootnoteWrapperCollection(Content); }
    }

    /// <summary>The endnotes of the main text, as <c>Document.Endnotes</c> in Word.</summary>
    public EndnoteWrapperCollection Endnotes
    {
        get { return new EndnoteWrapperCollection(Content); }
    }

    /// <summary>The list templates of the document, as <c>Document.ListTemplates</c> in Word.</summary>
    public ListTemplateWrapperCollection ListTemplates
    {
        get { return new ListTemplateWrapperCollection(this); }
    }

    /// <summary>The lists in the main text of the document, as <c>Document.Lists</c> in Word.</summary>
    public ListWrapperCollection Lists
    {
        get { return new ListWrapperCollection(this); }
    }

    /// <summary>The paragraphs of the main text that are in a list, as <c>Document.ListParagraphs</c> in Word.</summary>
    public ListParagraphWrapperCollection ListParagraphs
    {
        get
        {
            return new ListParagraphWrapperCollection(ListNumbering.GetItems(WordprocessingDocument.MainDocumentPart.Document.Body).Select(i => i.Paragraph).ToList());
        }
    }

    /// <summary>The hyperlinks in the main text of the document, as <c>Document.Hyperlinks</c> in Word.</summary>
    public HyperlinkWrapperCollection Hyperlinks
    {
        get { return new HyperlinkWrapperCollection(Content); }
    }

    /// <summary>The first range of each kind of story the document has, as <c>Document.StoryRanges</c> in Word.</summary>
    public StoryRangesWrapper StoryRanges
    {
        get { return new StoryRangesWrapper(this); }
    }

    /// <summary>Deletes every comment of the document, as <c>Document.DeleteAllComments</c> in Word.</summary>
    public void DeleteAllComments()
    {
        Comments.DeleteAll();
    }

    /// <summary>Sets the compatibility mode of the document, as <c>Document.SetCompatibilityMode</c> in Word: see <see cref="CompatibilityMode"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The mode is not 11, 12, 14 or 15.</exception>
    public void SetCompatibilityMode(int mode)
    {
        CompatibilityMode = mode;
    }

    /// <summary>
    /// Brings the document up to the latest format, compatibility mode 15, as <c>Document.Convert</c> in Word.
    /// </summary>
    /// <exception cref="InvalidOperationException">The document is in mode 15 already; Word makes the command unavailable.</exception>
    public void Convert()
    {
        if (CompatibilityMode == 15)
        {
            throw new InvalidOperationException("The document is in the latest format already.");
        }
        CompatibilityMode = 15;
    }
}
