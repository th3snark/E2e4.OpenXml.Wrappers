using System;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class WordprocessingDocumentWrapper : IDisposable
{
    private WordprocessingDocument _wordprocessingDocument;

    public WordprocessingDocumentWrapper(WordprocessingDocument wordprocessingDocument)
    {
        _wordprocessingDocument = wordprocessingDocument;
    }

    public WordprocessingDocument WordprocessingDocument
    {
        get { return _wordprocessingDocument; }
    }

    public static WordprocessingDocumentWrapper Open(string path)
    {
        return Open(path, false);
    }

    public static WordprocessingDocumentWrapper Open(string path, bool isEditable)
    {
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(path, isEditable);
        return new WordprocessingDocumentWrapper(wordprocessingDocument);
    }

    /// <summary>
    /// Opens a document from a stream. The stream must remain open for the life of the document and must be writable when
    /// <paramref name="isEditable"/> is true.
    /// </summary>
    public static WordprocessingDocumentWrapper Open(Stream stream, bool isEditable)
    {
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, isEditable);
        return new WordprocessingDocumentWrapper(wordprocessingDocument);
    }

    /// <summary>
    /// Creates a new document file. When the main document part is added it is initialised with an empty body (see
    /// <see cref="InitializeMainDocumentPart"/>) and, if requested, an empty comments part (see <see cref="InitializeCommentsPart"/>).
    /// </summary>
    public static WordprocessingDocumentWrapper Create(string path, bool addMainDocumentPart, bool addWordprocessingCommentsPart)
    {
        ValidateCreateArguments(addMainDocumentPart, addWordprocessingCommentsPart);
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        return InitializeNewDocument(wordprocessingDocument, addMainDocumentPart, addWordprocessingCommentsPart);
    }

    /// <summary>
    /// Creates a new document in a stream. See <see cref="Create(string, bool, bool)"/>.
    /// </summary>
    public static WordprocessingDocumentWrapper Create(Stream stream, bool addMainDocumentPart, bool addWordprocessingCommentsPart)
    {
        ValidateCreateArguments(addMainDocumentPart, addWordprocessingCommentsPart);
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        return InitializeNewDocument(wordprocessingDocument, addMainDocumentPart, addWordprocessingCommentsPart);
    }

    private static void ValidateCreateArguments(bool addMainDocumentPart, bool addWordprocessingCommentsPart)
    {
        if (addWordprocessingCommentsPart && !addMainDocumentPart)
        {
            throw new ArgumentException("A comments part can only be added together with a main document part.", nameof(addWordprocessingCommentsPart));
        }
    }

    private static WordprocessingDocumentWrapper InitializeNewDocument(WordprocessingDocument wordprocessingDocument, bool addMainDocumentPart, bool addWordprocessingCommentsPart)
    {
        WordprocessingDocumentWrapper wrapper = new WordprocessingDocumentWrapper(wordprocessingDocument);
        if (addMainDocumentPart)
        {
            wordprocessingDocument.AddMainDocumentPart();
            wrapper.InitializeMainDocumentPart();
            if (addWordprocessingCommentsPart)
            {
                wrapper.InitializeCommentsPart();
            }
        }
        return wrapper;
    }

    /// <summary>
    /// Saves any changes made to the document to its file or stream. Documents opened as editable also save when closed or disposed.
    /// </summary>
    public void Save()
    {
        _wordprocessingDocument.Save();
    }

    /// <summary>
    /// Saves a copy of the document to a new file and returns a wrapper over the copy. This document remains open and unchanged.
    /// </summary>
    public WordprocessingDocumentWrapper SaveAs(string path)
    {
        //Flush any changes held in the part trees so that the copy contains them.
        _wordprocessingDocument.Save();
        WordprocessingDocument copy = _wordprocessingDocument.Clone(path, true);

        //Clone does not carry over the core properties (title, author and so on), so copy them.
#pragma warning disable OOXML0001 // Only the name of the experimental interface is used; the properties themselves are not marked.
        IPackageProperties from = _wordprocessingDocument.PackageProperties;
        IPackageProperties to = copy.PackageProperties;
#pragma warning restore OOXML0001
        to.Title = from.Title;
        to.Subject = from.Subject;
        to.Creator = from.Creator;
        to.Keywords = from.Keywords;
        to.Description = from.Description;
        to.Category = from.Category;
        to.LastModifiedBy = from.LastModifiedBy;
        to.ContentStatus = from.ContentStatus;
        to.Language = from.Language;
        to.Version = from.Version;
        to.Revision = from.Revision;
        to.Created = from.Created;
        to.Modified = from.Modified;
        to.LastPrinted = from.LastPrinted;
        return new WordprocessingDocumentWrapper(copy);
    }

    public void Close()
    {
        Dispose();
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            _wordprocessingDocument.Dispose();
        }
    }

    public RangeWrapper Range
    {
        get { return new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document); }
    }

    public void InitializeMainDocumentPart()
    {
        //Initialize the main document part
        _wordprocessingDocument.MainDocumentPart.Document =
           new Document(
              new Body(
                 new Paragraph(),
                 new SectionProperties(
                    new PageSize() { Width = 11906, Height = 16838 },
                    new PageMargin() { Top = 1440, Right = 1800, Bottom = 1440, Left = 1800, Header = 708, Footer = 708, Gutter = 0 },
                    new Columns() { Space = "708" },
                    new DocGrid() { LinePitch = 360 }
                 )
              )
           );
    }

    public void InitializeCommentsPart()
    {
        MainDocumentPart mainDocumentPart = _wordprocessingDocument.MainDocumentPart;
        if (mainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Any() == false)
        {
            mainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
        }
        mainDocumentPart.WordprocessingCommentsPart.Comments = new Comments();
    }

    private static void CopyPartData(OpenXmlPart fromPart, OpenXmlPart toPart)
    {
        using (Stream stream = fromPart.GetStream(FileMode.Open))
        {
            toPart.FeedData(stream);
        }
    }

    public void CopyAllPartsTo(WordprocessingDocumentWrapper wordprocessingDocumentWrapper)
    {
        MainDocumentPart fromMainDocumentPart = _wordprocessingDocument.MainDocumentPart;
        MainDocumentPart toMainDocumentPart = wordprocessingDocumentWrapper.WordprocessingDocument.MainDocumentPart;

        ThemePart fromThemePart = fromMainDocumentPart.ThemePart;
        if (fromThemePart != null)
        {
            ThemePart toThemePart = toMainDocumentPart.AddNewPart<ThemePart>();
            CopyPartData(fromThemePart, toThemePart);
        }

        FontTablePart fromFontTablePart = fromMainDocumentPart.FontTablePart;
        if (fromFontTablePart != null)
        {
            FontTablePart toFontTablePart = toMainDocumentPart.AddNewPart<FontTablePart>();
            CopyPartData(fromFontTablePart, toFontTablePart);
        }

        StyleDefinitionsPart fromStyleDefinitionsPart = fromMainDocumentPart.StyleDefinitionsPart;
        if (fromStyleDefinitionsPart != null)
        {
            StyleDefinitionsPart toStyleDefinitionsPart = toMainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            CopyPartData(fromStyleDefinitionsPart, toStyleDefinitionsPart);
        }

        NumberingDefinitionsPart fromNumberingDefinitionsPart = fromMainDocumentPart.NumberingDefinitionsPart;
        if (fromNumberingDefinitionsPart != null)
        {
            NumberingDefinitionsPart toNumberingDefinitionsPart = toMainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            CopyPartData(fromNumberingDefinitionsPart, toNumberingDefinitionsPart);
        }

        FootnotesPart fromFootnotesPart = fromMainDocumentPart.FootnotesPart;
        if (fromFootnotesPart != null)
        {
            FootnotesPart toFootnotesPart = toMainDocumentPart.AddNewPart<FootnotesPart>();
            CopyPartData(fromFootnotesPart, toFootnotesPart);
        }

        EndnotesPart fromEndnotesPart = fromMainDocumentPart.EndnotesPart;
        if (fromEndnotesPart != null)
        {
            EndnotesPart toEndnotesPart = toMainDocumentPart.AddNewPart<EndnotesPart>();
            CopyPartData(fromEndnotesPart, toEndnotesPart);
        }

        DocumentSettingsPart fromDocumentSettingsPart = fromMainDocumentPart.DocumentSettingsPart;
        if (fromDocumentSettingsPart != null)
        {
            DocumentSettingsPart toDocumentSettingsPart = toMainDocumentPart.AddNewPart<DocumentSettingsPart>();
            CopyPartData(fromDocumentSettingsPart, toDocumentSettingsPart);
        }

        WebSettingsPart fromWebSettingsPart = fromMainDocumentPart.WebSettingsPart;
        if (fromWebSettingsPart != null)
        {
            WebSettingsPart toWebSettingsPart = toMainDocumentPart.AddNewPart<WebSettingsPart>();
            CopyPartData(fromWebSettingsPart, toWebSettingsPart);
        }

        //Copy the header parts to the target document with the same part id's
        foreach (HeaderPart fromHeaderPart in fromMainDocumentPart.HeaderParts)
        {
            string fromHeaderPartId = fromMainDocumentPart.GetIdOfPart(fromHeaderPart);
            HeaderPart toHeaderPart = toMainDocumentPart.AddNewPart<HeaderPart>(fromHeaderPartId);
            CopyPartData(fromHeaderPart, toHeaderPart);
        }

        //Copy the footer parts to the target document with the same part id's
        foreach (FooterPart fromFooterPart in fromMainDocumentPart.FooterParts)
        {
            string fromFooterPartId = fromMainDocumentPart.GetIdOfPart(fromFooterPart);
            FooterPart toFooterPart = toMainDocumentPart.AddNewPart<FooterPart>(fromFooterPartId);
            CopyPartData(fromFooterPart, toFooterPart);
        }
    }

    public void CopyLayoutTo(WordprocessingDocumentWrapper wordprocessingDocumentWrapper)
    {
        MainDocumentPart toMainDocumentPart = wordprocessingDocumentWrapper.WordprocessingDocument.MainDocumentPart;

        //Delete the existing header and footer parts from the target document
        if (toMainDocumentPart.GetPartsOfType<HeaderPart>().Any())
        {
            toMainDocumentPart.DeleteParts(toMainDocumentPart.HeaderParts);
        }
        if (toMainDocumentPart.GetPartsOfType<FooterPart>().Any())
        {
            toMainDocumentPart.DeleteParts(toMainDocumentPart.FooterParts);
        }

        //Replace the section properties in the target document with that of the source document
        MainDocumentPart fromMainDocumentPart = _wordprocessingDocument.MainDocumentPart;
        SectionProperties fromSectionProperties = fromMainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
        SectionProperties toSectionProperties = toMainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
        if (toSectionProperties != null)
        {
            toSectionProperties.Parent.RemoveChild(toSectionProperties);
        }

        //Create a cloned copy of the source section properties
        SectionProperties cloneOfFromSectionProperties = null;
        if (fromSectionProperties != null)
        {
            cloneOfFromSectionProperties = toMainDocumentPart.Document.Body.AppendChild((SectionProperties)fromSectionProperties.Clone());
        }

        //If we have a cloned copy then process each header and fix up the section properties header reference id's
        if (cloneOfFromSectionProperties != null)
        {
            foreach (HeaderReference headerReference in cloneOfFromSectionProperties.Elements<HeaderReference>())
            {
                string fromId = headerReference.Id;
                HeaderPart fromHeaderPart = fromMainDocumentPart.HeaderParts.Where(p => fromMainDocumentPart.GetIdOfPart(p) == fromId).FirstOrDefault();

                //Add a new header part to the target document
                HeaderPart toHeaderPart = toMainDocumentPart.AddNewPart<HeaderPart>();
                string toId = toMainDocumentPart.GetIdOfPart(toHeaderPart);

                //Feed the targer header part with the source header part
                CopyPartData(fromHeaderPart, toHeaderPart);

                //Update the header reference with the new id
                headerReference.Id = toId;
            }

            //If we have a cloned copy then process each footer and fix up the section properties footer reference id's
            foreach (FooterReference footerReference in cloneOfFromSectionProperties.Elements<FooterReference>())
            {
                string fromId = footerReference.Id;
                FooterPart fromFooterPart = fromMainDocumentPart.FooterParts.Where(p => fromMainDocumentPart.GetIdOfPart(p) == fromId).FirstOrDefault();

                //Add a new footer part to the target document
                FooterPart toFooterPart = toMainDocumentPart.AddNewPart<FooterPart>();
                string toId = toMainDocumentPart.GetIdOfPart(toFooterPart);

                //Feed the targer footer part with the source footer part
                CopyPartData(fromFooterPart, toFooterPart);

                //Update the footer reference with the new id
                footerReference.Id = toId;
            }
        }
    }

    public TableWrapperCollection Tables
    {
        get { return new TableWrapperCollection(new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document)); }
    }

    public BookmarkWrapperCollection Bookmarks
    {
        get { return new BookmarkWrapperCollection(new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document)); }
    }

    public ParagraphWrapperCollection Paragraphs
    {
        get { return new ParagraphWrapperCollection(new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document)); }
    }

    /// <summary>
    /// The sections of the document. A document always has at least one.
    /// </summary>
    public SectionWrapperCollection Sections
    {
        get { return new SectionWrapperCollection(_wordprocessingDocument.MainDocumentPart.Document.Body); }
    }

    /// <summary>
    /// The styles of the document.
    /// </summary>
    public StyleWrapperCollection Styles
    {
        get { return new StyleWrapperCollection(_wordprocessingDocument.MainDocumentPart); }
    }

    /// <summary>
    /// The built-in properties of the document (title, author and so on).
    /// </summary>
    public DocumentPropertiesWrapper BuiltInDocumentProperties
    {
        get { return new DocumentPropertiesWrapper(_wordprocessingDocument); }
    }

    public CommentWrapperCollection Comments
    {
        get { return new CommentWrapperCollection(new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document)); }
    }
}
