using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class WordprocessingDocumentWrapper
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

    public static WordprocessingDocumentWrapper Create(string path, bool addMainDocumentPart, bool addWordprocessingCommentsPart)
    {
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        if (addMainDocumentPart)
        {
            MainDocumentPart mainDocumentPart = wordprocessingDocument.AddMainDocumentPart();
            if (addWordprocessingCommentsPart)
            {
                WordprocessingCommentsPart wordprocessingCommentsPart = mainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            }
        }
        return new WordprocessingDocumentWrapper(wordprocessingDocument);
    }

    public void Close()
    {
        _wordprocessingDocument.Dispose();
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

    public CommentWrapperCollection Comments
    {
        get { return new CommentWrapperCollection(new RangeWrapper(_wordprocessingDocument.MainDocumentPart.Document)); }
    }
}
