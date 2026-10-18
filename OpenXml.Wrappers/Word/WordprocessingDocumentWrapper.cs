using System;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class WordprocessingDocumentWrapper : IDisposable
{
    private WordprocessingDocument _wordprocessingDocument;
    private bool _skipSave;
    private bool _disposed;

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
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(path, isEditable, new OpenSettings() { AutoSave = false });
        FieldUpdater.RememberPath(wordprocessingDocument, path);
        return new WordprocessingDocumentWrapper(wordprocessingDocument);
    }

    /// <summary>
    /// Opens a document from a stream. The stream must remain open for the life of the document and must be writable when
    /// <paramref name="isEditable"/> is true.
    /// </summary>
    public static WordprocessingDocumentWrapper Open(Stream stream, bool isEditable)
    {
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, isEditable, new OpenSettings() { AutoSave = false });
        return new WordprocessingDocumentWrapper(wordprocessingDocument);
    }

    /// <summary>
    /// Creates a new document file. When the main document part is added it is initialised with an empty body (see
    /// <see cref="InitializeMainDocumentPart"/>) and, if requested, an empty comments part (see <see cref="InitializeCommentsPart"/>).
    /// </summary>
    public static WordprocessingDocumentWrapper Create(string path, bool addMainDocumentPart, bool addWordprocessingCommentsPart)
    {
        ValidateCreateArguments(addMainDocumentPart, addWordprocessingCommentsPart);
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document, false);
        FieldUpdater.RememberPath(wordprocessingDocument, path);
        return InitializeNewDocument(wordprocessingDocument, addMainDocumentPart, addWordprocessingCommentsPart);
    }

    /// <summary>
    /// Creates a new document in a stream. See <see cref="Create(string, bool, bool)"/>.
    /// </summary>
    public static WordprocessingDocumentWrapper Create(Stream stream, bool addMainDocumentPart, bool addWordprocessingCommentsPart)
    {
        ValidateCreateArguments(addMainDocumentPart, addWordprocessingCommentsPart);
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, false);
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

    //Saves the package after the headers and footers that were only read and are still empty have been taken out.
    private void SaveChanges()
    {
        HeaderFooterWrapper.RemoveLatentStories(_wordprocessingDocument);
        _wordprocessingDocument.Save();
    }

    /// <summary>
    /// Saves any changes made to the document to its file or stream. Documents opened as editable also save when closed or disposed.
    /// </summary>
    public void Save()
    {
        SaveChanges();
    }

    /// <summary>
    /// Saves a copy of the document to a new file and returns a wrapper over the copy. This document remains open and unchanged.
    /// </summary>
    public WordprocessingDocumentWrapper SaveAs(string path)
    {
        //Flush any changes held in the part trees so that the copy contains them.
        SaveChanges();
        WordprocessingDocument copy = _wordprocessingDocument.Clone(path, true, new OpenSettings() { AutoSave = false });

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
        FieldUpdater.RememberPath(copy, path);
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
        if (disposing && !_disposed)
        {
            _disposed = true;
            //A document opened by this library does not save itself, so that Close can leave the changes out.
            if (!_skipSave && !_wordprocessingDocument.AutoSave && _wordprocessingDocument.FileOpenAccess != FileAccess.Read)
            {
                SaveChanges();
            }
            _wordprocessingDocument.Dispose();
        }
    }

    /// <summary>
    /// The whole content of the document, as <c>Document.Content</c> in Word.
    /// </summary>
    public RangeWrapper Content
    {
        get { return new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document); }
    }

    /// <summary>
    /// The positions of the main text of the document, worked out once so that the positions of many ranges can be read at once.
    /// </summary>
    public PositionIndex CreatePositionIndex()
    {
        return Content.CreatePositionIndex();
    }

    /// <summary>
    /// The span of characters from <paramref name="start"/> up to <paramref name="end"/> in the main text of the document, as
    /// <c>Document.Range(Start, End)</c> in Word. See <see cref="RangeWrapper.Start"/> for how positions are counted.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The positions are not in order or are beyond the end of the text.</exception>
    public RangeWrapper Range(int start, int end)
    {
        RangeWrapper range = Content;
        range.SetRange(start, end);
        return range;
    }

    /// <summary>
    /// The whole main text of the document, as <c>Document.Range</c> without arguments in Word. The same as <see cref="Content"/>.
    /// </summary>
    public RangeWrapper Range()
    {
        return Content;
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

        //A document without compatibility settings is opened by Word in Compatibility Mode (as a Word 2007 document), so say that it is a
        //current one.
        MainDocumentPart mainDocumentPart = _wordprocessingDocument.MainDocumentPart;
        if (mainDocumentPart.DocumentSettingsPart == null)
        {
            DocumentSettingsPart settingsPart = mainDocumentPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                new Compatibility(
                    new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" }));
        }
    }

    public void InitializeCommentsPart()
    {
        MainDocumentPart mainDocumentPart = _wordprocessingDocument.MainDocumentPart;
        if (mainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Any() == false)
        {
            mainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
        }
        if (mainDocumentPart.WordprocessingCommentsPart.Comments == null)
        {
            mainDocumentPart.WordprocessingCommentsPart.Comments = new Comments();
        }
    }

    private static void CopyPartData(OpenXmlPart fromPart, OpenXmlPart toPart)
    {
        using (Stream stream = fromPart.GetStream(FileMode.Open))
        {
            toPart.FeedData(stream);
        }
    }

    //Copies a part of the main document part into the target, replacing the part of that kind that the target already has.
    private static void CopyPart<T>(T fromPart, MainDocumentPart toMainDocumentPart, T existingPart) where T : OpenXmlPart, IFixedContentTypePart
    {
        if (fromPart == null)
        {
            return;
        }
        if (existingPart != null)
        {
            toMainDocumentPart.DeletePart(existingPart);
        }
        CopyPartData(fromPart, toMainDocumentPart.AddNewPart<T>());
    }

    /// <summary>
    /// Copies the theme, font table, styles, numbering, footnotes, endnotes, settings, web settings, headers and footers of this document
    /// into the other, replacing the ones that it already has.
    /// </summary>
    public void CopyAllPartsTo(WordprocessingDocumentWrapper wordprocessingDocumentWrapper)
    {
        MainDocumentPart fromMainDocumentPart = _wordprocessingDocument.MainDocumentPart;
        MainDocumentPart toMainDocumentPart = wordprocessingDocumentWrapper.WordprocessingDocument.MainDocumentPart;

        CopyPart(fromMainDocumentPart.ThemePart, toMainDocumentPart, toMainDocumentPart.ThemePart);
        CopyPart(fromMainDocumentPart.FontTablePart, toMainDocumentPart, toMainDocumentPart.FontTablePart);
        CopyPart(fromMainDocumentPart.StyleDefinitionsPart, toMainDocumentPart, toMainDocumentPart.StyleDefinitionsPart);
        CopyPart(fromMainDocumentPart.NumberingDefinitionsPart, toMainDocumentPart, toMainDocumentPart.NumberingDefinitionsPart);
        CopyPart(fromMainDocumentPart.FootnotesPart, toMainDocumentPart, toMainDocumentPart.FootnotesPart);
        CopyPart(fromMainDocumentPart.EndnotesPart, toMainDocumentPart, toMainDocumentPart.EndnotesPart);
        CopyPart(fromMainDocumentPart.DocumentSettingsPart, toMainDocumentPart, toMainDocumentPart.DocumentSettingsPart);
        CopyPart(fromMainDocumentPart.WebSettingsPart, toMainDocumentPart, toMainDocumentPart.WebSettingsPart);

        //Copy the header and footer parts to the target document with the same part id's, replacing the ones it has
        if (toMainDocumentPart.HeaderParts.Any())
        {
            toMainDocumentPart.DeleteParts(toMainDocumentPart.HeaderParts.ToList());
        }
        if (toMainDocumentPart.FooterParts.Any())
        {
            toMainDocumentPart.DeleteParts(toMainDocumentPart.FooterParts.ToList());
        }
        foreach (HeaderPart fromHeaderPart in fromMainDocumentPart.HeaderParts)
        {
            string fromHeaderPartId = fromMainDocumentPart.GetIdOfPart(fromHeaderPart);
            HeaderPart toHeaderPart = toMainDocumentPart.AddNewPart<HeaderPart>(fromHeaderPartId);
            CopyPartData(fromHeaderPart, toHeaderPart);
        }
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
            //References to parts that the source does not have are dropped.
            List<HeaderReference> missingHeaders = new List<HeaderReference>();
            List<FooterReference> missingFooters = new List<FooterReference>();
            foreach (HeaderReference headerReference in cloneOfFromSectionProperties.Elements<HeaderReference>().ToList())
            {
                string fromId = headerReference.Id;
                HeaderPart fromHeaderPart = fromMainDocumentPart.HeaderParts.Where(p => fromMainDocumentPart.GetIdOfPart(p) == fromId).FirstOrDefault();
                if (fromHeaderPart == null)
                {
                    missingHeaders.Add(headerReference);
                    continue;
                }

                //Add a new header part to the target document
                HeaderPart toHeaderPart = toMainDocumentPart.AddNewPart<HeaderPart>();
                string toId = toMainDocumentPart.GetIdOfPart(toHeaderPart);

                //Feed the targer header part with the source header part
                CopyPartData(fromHeaderPart, toHeaderPart);

                //Update the header reference with the new id
                headerReference.Id = toId;
            }

            //If we have a cloned copy then process each footer and fix up the section properties footer reference id's
            foreach (FooterReference footerReference in cloneOfFromSectionProperties.Elements<FooterReference>().ToList())
            {
                string fromId = footerReference.Id;
                FooterPart fromFooterPart = fromMainDocumentPart.FooterParts.Where(p => fromMainDocumentPart.GetIdOfPart(p) == fromId).FirstOrDefault();
                if (fromFooterPart == null)
                {
                    missingFooters.Add(footerReference);
                    continue;
                }

                //Add a new footer part to the target document
                FooterPart toFooterPart = toMainDocumentPart.AddNewPart<FooterPart>();
                string toId = toMainDocumentPart.GetIdOfPart(toFooterPart);

                //Feed the targer footer part with the source footer part
                CopyPartData(fromFooterPart, toFooterPart);

                //Update the footer reference with the new id
                footerReference.Id = toId;
            }

            foreach (HeaderReference missing in missingHeaders)
            {
                missing.Remove();
            }
            foreach (FooterReference missing in missingFooters)
            {
                missing.Remove();
            }
        }
    }

    public TableWrapperCollection Tables
    {
        get { return new TableWrapperCollection(new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document)); }
    }

    public FieldWrapperCollection Fields
    {
        get { return new FieldWrapperCollection(new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document)); }
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

    /// <summary>
    /// The custom properties of the document.
    /// </summary>
    public CustomDocumentPropertyCollection CustomDocumentProperties
    {
        get { return new CustomDocumentPropertyCollection(_wordprocessingDocument); }
    }

    public CommentWrapperCollection Comments
    {
        get { return new CommentWrapperCollection(new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document)); }
    }
}
