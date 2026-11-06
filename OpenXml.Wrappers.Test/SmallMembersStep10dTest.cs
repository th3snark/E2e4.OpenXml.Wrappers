using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Comment.Reference and DeleteRecursively, Field.Kind, Fields.Unlink, Bookmark.Column, Section.ProtectedForForms and Sections.PageSetup. Values and
/// markup are Word 16's (the scenarios of MarkupConformanceTest and the field kinds measured in Word).
/// </summary>
public class SmallMembersStep10dTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public SmallMembersStep10dTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ACommentsReferenceIsTheMarkAfterItsScope()
    {
        _document.Content.InsertAfter("hello");
        CommentWrapper comment = _document.Comments.Add(_document.Range(0, 5), "note");

        Assert.Equal(string.Empty, comment.Reference.Text);
        Assert.Equal(comment.Scope.End, comment.Reference.Start);
    }

    [Fact]
    public void DeleteRecursivelyTakesTheRepliesToo()
    {
        _document.Content.InsertAfter("hello");
        CommentWrapper comment = _document.Comments.Add(_document.Range(0, 5), "note");
        comment.Replies.Add("reply");
        Assert.Equal(2, _document.Comments.Count);

        comment.DeleteRecursively();

        Assert.Equal(0, _document.Comments.Count);
    }

    [Theory]
    [InlineData(FieldType.MacroButton, FieldKind.Hot)]
    [InlineData(FieldType.Symbol, FieldKind.Hot)]
    [InlineData(FieldType.AutoNum, FieldKind.Hot)]
    [InlineData(FieldType.Page, FieldKind.Warm)]
    [InlineData(FieldType.Date, FieldKind.Warm)]
    [InlineData(FieldType.Hyperlink, FieldKind.Warm)]
    [InlineData(FieldType.IndexEntry, FieldKind.Cold)]
    [InlineData(FieldType.TocEntry, FieldKind.Cold)]
    [InlineData(FieldType.Private, FieldKind.Cold)]
    public void FieldKindsAreWords(FieldType type, FieldKind kind)
    {
        string code = type == FieldType.MacroButton ? "MACROBUTTON NoMacro click"
            : type == FieldType.Symbol ? "SYMBOL 65"
            : type == FieldType.AutoNum ? "AUTONUM"
            : type == FieldType.Page ? "PAGE"
            : type == FieldType.Date ? "DATE"
            : type == FieldType.Hyperlink ? "HYPERLINK \"http://example.com\""
            : type == FieldType.IndexEntry ? "XE \"a\""
            : type == FieldType.TocEntry ? "TC \"a\"" : "PRIVATE x";
        _document.Content.InsertAfter("text");

        FieldWrapper field = _document.Fields.Add(_document.Range(0, 0), FieldType.Empty, code, false);

        Assert.Equal(type, field.Type);
        Assert.Equal(kind, field.Kind);
    }

    [Fact]
    public void UnlinkingTheFieldsLeavesTheirResults()
    {
        _document.Content.InsertAfter("text");
        _document.Fields.Add(_document.Range(0, 0), FieldType.Empty, "SEQ a", false);
        _document.Fields.Add(_document.Range(0, 0), FieldType.Empty, "SEQ b", false);
        Assert.Equal(2, _document.Fields.Count);

        _document.Fields.Unlink();

        Assert.Equal(0, _document.Fields.Count);
        Assert.Contains("text", _document.Content.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ABookmarkOverACellIsAColumnBookmark()
    {
        _document.Tables.Add(_document.Range(0, 0), 3, 3);
        _document.Bookmarks.Add("Plain", _document.Range(0, 0));
        _document.Bookmarks.Add("Cell", _document.Range(0, 0));
        //Word writes the columns of a bookmark made over cells; the library reads them.
        _document.WordprocessingDocument.MainDocumentPart.Document.Descendants<BookmarkStart>().Single(b => b.Name.Value == "Cell").ColumnFirst = 1;

        Assert.False(_document.Bookmarks["Plain"].Column);
        Assert.True(_document.Bookmarks["Cell"].Column);
    }

    [Fact]
    public void SectionsAreProtectedForFormsUnlessSaidOtherwise()
    {
        _document.Content.InsertAfter("a");
        _document.Sections.Add();
        Assert.True(_document.Sections[0].ProtectedForForms);
        Assert.True(_document.Sections[1].ProtectedForForms);

        _document.Sections[0].ProtectedForForms = false;

        Assert.False(_document.Sections[0].ProtectedForForms);
        Assert.True(_document.Sections[1].ProtectedForForms);
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.Document.Descendants<FormProtection>());

        _document.Sections[0].ProtectedForForms = true;

        Assert.True(_document.Sections[0].ProtectedForForms);
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.Document.Descendants<FormProtection>());
    }

    [Fact]
    public void TheSectionsPageSetupAppliesToEveryOne()
    {
        _document.Content.InsertAfter("a");
        _document.Sections.Add();

        _document.Sections.PageSetup.LeftMargin = 50;

        Assert.Equal(50, _document.Sections[0].PageSetup.LeftMargin);
        Assert.Equal(50, _document.Sections[1].PageSetup.LeftMargin);
    }
}
