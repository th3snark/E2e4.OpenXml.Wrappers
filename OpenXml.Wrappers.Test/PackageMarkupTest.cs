using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// What goes between the parts of a document: ids that must be unique across them, and what a copied part refers to.
/// </summary>
public class PackageMarkupTest : IDisposable
{
    private readonly MemoryStream _stream = new MemoryStream();
    private readonly WordprocessingDocumentWrapper _document;

    public PackageMarkupTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_stream, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        _stream.Dispose();
        GC.SuppressFinalize(this);
    }

    private MainDocumentPart Main
    {
        get { return _document.WordprocessingDocument.MainDocumentPart; }
    }

    private void AssertClean()
    {
        _document.Save();
        foreach (string check in DocumentIntegrity.All)
        {
            Assert.Empty(DocumentIntegrity.Check(_document.WordprocessingDocument, check));
        }
    }

    [Fact]
    public void BookmarkIdsAreUniqueAcrossTheParts()
    {
        _document.Content.InsertAfter("Body text");
        _document.Bookmarks.Add("InBody", _document.Range(0, 4));
        HeaderFooterWrapper header = _document.Sections[0].Headers[HeaderFooterIndex.Primary];
        header.Range.Text = "Header text";

        RangeWrapper paragraph = new RangeWrapper(Main.HeaderParts.Single().Header.Elements<Paragraph>().Single());
        new BookmarkWrapperCollection(paragraph).Add("InHeader", paragraph);

        string bodyId = Main.Document.Descendants<BookmarkStart>().Single().Id.Value;
        string headerId = Main.HeaderParts.Single().Header.Descendants<BookmarkStart>().Single().Id.Value;
        Assert.NotEqual(bodyId, headerId);
        AssertClean();
    }

    [Fact]
    public void AHeaderThatIsUnlinkedKeepsWhatItRefersTo()
    {
        HeaderFooterWrapper first = _document.Sections[0].Headers[HeaderFooterIndex.Primary];
        first.Range.Text = "Header";
        HeaderPart part = Main.HeaderParts.Single();
        string relationship = part.AddHyperlinkRelationship(new Uri("https://example.com/"), true).Id;
        part.Header.Elements<Paragraph>().First().AppendChild(new Hyperlink(new Run(new Text("link"))) { Id = relationship });
        part.Header.Elements<Paragraph>().First().AppendChild(new BookmarkStart() { Id = "9", Name = "Mark" });
        part.Header.Elements<Paragraph>().First().AppendChild(new BookmarkEnd() { Id = "9" });
        _document.Sections.Add();

        _document.Sections[1].Headers[HeaderFooterIndex.Primary].LinkToPrevious = false;

        HeaderPart copy = Main.HeaderParts.Single(p => p != part);
        Assert.Equal(relationship, copy.HyperlinkRelationships.Single().Id);
        Assert.Empty(copy.Header.Descendants<BookmarkStart>());
        Assert.Single(part.Header.Descendants<BookmarkStart>());
        AssertClean();
    }

    [Fact]
    public void ACopiedRowLeavesItsBookmarksAndCommentsBehind()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 1);
        Run run = new Run(new Text("Cell"));
        Paragraph paragraph = table.Table.Descendants<Paragraph>().Single();
        paragraph.ParagraphId = "12345678";
        paragraph.AppendChild(run);
        _document.Bookmarks.Add("Mark", new RangeWrapper(paragraph));
        _document.Comments.Add(new RangeWrapper(paragraph), "A comment");

        table.AppendRow(table.Rows[0].Copy());

        Assert.Equal(2, table.Rows.Count);
        Assert.Equal("Cell", table.Rows[1].TableRow.Descendants<Paragraph>().Single().InnerText);
        Assert.Single(Main.Document.Descendants<BookmarkStart>());
        Assert.Single(Main.Document.Descendants<CommentReference>());
        Assert.Null(table.Rows[1].TableRow.Descendants<Paragraph>().Single().ParagraphId);
        AssertClean();
    }

    [Fact]
    public void DeletingACellOrColumnTakesItsAnchorsAway()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 3);
        foreach (int column in new[] { 0, 1 })
        {
            RangeWrapper range = table.Cell(0, column).Range;
            range.InsertBefore("x");
            _document.Comments.Add(range, "A comment");
            _document.Bookmarks.Add("Mark" + column.ToString(System.Globalization.CultureInfo.InvariantCulture), range);
        }

        table.Columns[1].Delete();
        Assert.Single(Main.Document.Descendants<BookmarkStart>());
        Assert.Single(Main.WordprocessingCommentsPart.Comments.Elements<Comment>());

        table.Cell(0, 0).Delete();
        Assert.Empty(Main.Document.Descendants<BookmarkStart>());
        Assert.Empty(Main.WordprocessingCommentsPart.Comments.Elements<Comment>());
        Assert.Empty(Main.WordprocessingCommentsIdsPart.CommentsIds.ChildElements);
    }

    [Fact]
    public void CopyingLayoutCopiesWhatHeadersReferToAndLeavesNoStaleReferences()
    {
        HeaderFooterWrapper header = _document.Sections[0].Headers[HeaderFooterIndex.Primary];
        header.Range.Text = "Source header";
        string relationship = Main.HeaderParts.Single().AddHyperlinkRelationship(new Uri("https://example.com/"), true).Id;
        _document.Save();

        using (MemoryStream targetStream = new MemoryStream())
        using (WordprocessingDocumentWrapper target = WordprocessingDocumentWrapper.Create(targetStream, true, false))
        {
            target.Content.InsertAfter("First");
            target.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Old header";
            target.Sections.Add();

            _document.CopyLayoutTo(target);

            MainDocumentPart main = target.WordprocessingDocument.MainDocumentPart;
            HeaderPart copied = main.HeaderParts.Single();
            Assert.Equal("Source header", copied.Header.InnerText);
            Assert.Equal(relationship, copied.HyperlinkRelationships.Single().Id);
            string id = main.GetIdOfPart(copied);
            Assert.All(main.Document.Descendants<HeaderReference>(), r => Assert.Equal(id, r.Id.Value));
            target.Save();
            Assert.Empty(DocumentIntegrity.Check(target.WordprocessingDocument, DocumentIntegrity.Relationships));
        }
    }

    [Fact]
    public void ACommentHasItsEntryInEveryPartAndTheyGoWithIt()
    {
        _document.Content.InsertAfter("Commented text");

        CommentWrapper comment = _document.Comments.Add(_document.Range(0, 9), "A comment");
        Assert.Single(Main.WordprocessingCommentsExPart.CommentsEx.ChildElements);
        Assert.Single(Main.WordprocessingCommentsIdsPart.CommentsIds.ChildElements);
        Assert.Single(Main.WordCommentsExtensiblePart.CommentsExtensible.ChildElements);
        Assert.Single(Main.WordprocessingPeoplePart.People.ChildElements);
        AssertClean();

        comment.Delete();
        Assert.Empty(Main.WordprocessingCommentsExPart.CommentsEx.ChildElements);
        Assert.Empty(Main.WordprocessingCommentsIdsPart.CommentsIds.ChildElements);
        Assert.Empty(Main.WordCommentsExtensiblePart.CommentsExtensible.ChildElements);
    }

    [Fact]
    public void AChangedAuthorIsListedAndTheOldOneIsNot()
    {
        _document.Content.InsertAfter("Text");
        CommentWrapper comment = _document.Comments.Add(_document.Range(0, 4), "A comment");

        comment.Author = "Someone Else";

        Assert.Equal("Someone Else", Main.WordprocessingPeoplePart.People.ChildElements.Cast<DocumentFormat.OpenXml.Office2013.Word.Person>().Single().Author.Value);
    }

    [Fact]
    public void ACommentDateIsLocalAndItsUtcTimeIsKeptBesideIt()
    {
        _document.Content.InsertAfter("Text");
        CommentWrapper comment = _document.Comments.Add(_document.Range(0, 4), "A comment");
        DateTime utc = new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc);

        comment.Date = utc;

        Assert.Equal(DateTime.SpecifyKind(utc.ToLocalTime(), DateTimeKind.Unspecified), comment.Date);
        Assert.Equal(utc, Main.WordCommentsExtensiblePart.CommentsExtensible.ChildElements.Cast<DocumentFormat.OpenXml.Office2021.Word.CommentsExt.CommentExtensible>().Single().DateUtc.Value.ToUniversalTime());
    }

    [Fact]
    public void AnEmptyCommentHasNoTextRun()
    {
        _document.Comments.Add(_document.Range(0, 0), string.Empty);

        Comment comment = Main.WordprocessingCommentsPart.Comments.Elements<Comment>().Single();
        Assert.Empty(comment.Descendants<Text>());
        Assert.Empty(Main.Document.Descendants<Text>());
    }

    [Fact]
    public void APasswordThatProtectsADocumentAlsoUnprotectsIt()
    {
        _document.Protect(ProtectionType.AllowOnlyReading, "password");
        DocumentProtection protection = Main.DocumentSettingsPart.Settings.GetFirstChild<DocumentProtection>();
        Assert.NotNull(protection.Hash);
        Assert.Null(protection.HashValue);

        Assert.Throws<UnauthorizedAccessException>(() => _document.Unprotect("wrong"));
        _document.Unprotect("password");
        Assert.Equal(ProtectionType.NoProtection, _document.ProtectionType);
    }

    [Fact]
    public void ANewSectionInheritsItsHeadersAndTheBreakCopiesTheSectionItIsIn()
    {
        _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Header";
        _document.Content.InsertAfter("First");
        _document.Sections[0].PageSetup.LeftMargin = 50;

        _document.Sections.Add();

        SectionProperties[] sections = Main.Document.Descendants<SectionProperties>().ToArray();
        Assert.Single(sections[0].Elements<HeaderReference>());
        Assert.Empty(sections[1].Elements<HeaderReference>());
        Assert.Equal(50, _document.Sections[0].PageSetup.LeftMargin);
        Assert.Equal(50, _document.Sections[1].PageSetup.LeftMargin);
        AssertClean();
    }

    [Fact]
    public void ABreakInAnEarlierSectionCopiesThatSectionNotTheLast()
    {
        _document.Content.InsertAfter("One");
        _document.Content.InsertParagraphAfter();
        _document.Content.InsertAfter("Two");
        _document.Content.InsertParagraphAfter();
        _document.Content.InsertAfter("Three");
        new ParagraphWrapper(_document.Paragraphs[1].Paragraph).InsertSectionBreak();
        _document.Sections[0].PageSetup.LeftMargin = 50;
        _document.Sections[1].PageSetup.LeftMargin = 80;

        new ParagraphWrapper(_document.Paragraphs[0].Paragraph).InsertSectionBreak();

        Assert.Equal(3, _document.Sections.Count);
        Assert.Equal(50, _document.Sections[0].PageSetup.LeftMargin);
        Assert.Equal(50, _document.Sections[1].PageSetup.LeftMargin);
        Assert.Equal(80, _document.Sections[2].PageSetup.LeftMargin);
    }

    [Fact]
    public void ALandscapeSectionThatGoesBackToPortraitSaysNothing()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;

        setup.Orientation = PageOrientation.Landscape;
        Assert.Equal(PageOrientationValues.Landscape, Main.Document.Body.Elements<SectionProperties>().Single().GetFirstChild<PageSize>().Orient.Value);

        setup.Orientation = PageOrientation.Portrait;
        Assert.Null(Main.Document.Body.Elements<SectionProperties>().Single().GetFirstChild<PageSize>().Orient);
        Assert.Equal(PageOrientation.Portrait, setup.Orientation);
    }
}
