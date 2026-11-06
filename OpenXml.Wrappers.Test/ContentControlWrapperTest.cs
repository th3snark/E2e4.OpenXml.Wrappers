using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Content controls. The markup is Word 16's (the content-control-* scenarios of MarkupConformanceTest); the values of the members are those Word
/// gave for the same calls.
/// </summary>
public class ContentControlWrapperTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ContentControlWrapperTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Hello world again");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        System.Collections.Generic.IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    [Theory]
    [InlineData(ContentControlType.RichText)]
    [InlineData(ContentControlType.Text)]
    [InlineData(ContentControlType.ComboBox)]
    [InlineData(ContentControlType.DropdownList)]
    [InlineData(ContentControlType.Date)]
    public void AControlAroundTextHoldsTheText(ContentControlType type)
    {
        ContentControlWrapper control = _document.ContentControls.Add(type, _document.Range(6, 11));

        Assert.Equal(type, control.Type);
        Assert.Equal("world", control.Range.Text);
        Assert.Equal(7, control.Range.Start);
        Assert.Equal(12, control.Range.End);
        Assert.False(control.ShowingPlaceholderText);
        Assert.Equal(string.Empty, control.Title);
        Assert.Equal(string.Empty, control.Tag);
        Assert.Equal("Hello world again", _document.Content.Text.TrimEnd('\r'));
        AssertValid();
    }

    [Theory]
    [InlineData(ContentControlType.Text, "Click or tap here to enter text.")]
    [InlineData(ContentControlType.DropdownList, "Choose an item.")]
    [InlineData(ContentControlType.Date, "Click or tap to enter a date.")]
    public void AControlAtAPointShowsItsPlaceholder(ContentControlType type, string placeholder)
    {
        ContentControlWrapper control = _document.ContentControls.Add(type, _document.Range(5, 5));

        Assert.True(control.ShowingPlaceholderText);
        Assert.Equal(placeholder, control.Range.Text);
        Assert.NotNull(_document.WordprocessingDocument.MainDocumentPart.GlossaryDocumentPart);
        AssertValid();
    }

    [Fact]
    public void ACheckBoxHasNoPlaceholderAndNeedsNoGlossary()
    {
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.CheckBox, _document.Range(6, 6));

        Assert.Equal(ContentControlType.CheckBox, control.Type);
        Assert.Equal("☐", control.Range.Text);
        Assert.Null(_document.WordprocessingDocument.MainDocumentPart.GlossaryDocumentPart);
        Assert.Throws<NotSupportedException>(() => _document.ContentControls.Add(ContentControlType.CheckBox, _document.Range(0, 5)));
        AssertValid();
    }

    [Fact]
    public void ThePropertiesReadBackInTheOrderWordWritesThem()
    {
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.RichText, _document.Range(6, 11));

        control.Title = "My title";
        control.Tag = "my-tag";
        control.LockContentControl = true;
        control.LockContents = true;
        control.Appearance = ContentControlAppearance.Hidden;

        Assert.Equal("My title", control.Title);
        Assert.Equal("my-tag", control.Tag);
        Assert.True(control.LockContentControl);
        Assert.True(control.LockContents);
        Assert.Equal(ContentControlAppearance.Hidden, control.Appearance);

        control.LockContents = false;
        control.Title = null;

        Assert.True(control.LockContentControl);
        Assert.False(control.LockContents);
        Assert.Equal(string.Empty, control.Title);
        AssertValid();
    }

    [Fact]
    public void TheIdIsANumberWrittenAsText()
    {
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(6, 11));

        Assert.True(uint.TryParse(control.ID, out uint id));
        Assert.True(id > 0);
    }

    [Fact]
    public void ControlsAreFoundByTagAndTitleAndInRanges()
    {
        ContentControlWrapper first = _document.ContentControls.Add(ContentControlType.Text, _document.Range(0, 5));
        first.Tag = "a";
        ContentControlWrapper second = _document.ContentControls.Add(ContentControlType.Text, _document.Range(_document.ContentControls[0].Range.End + 7, _document.ContentControls[0].Range.End + 11));
        second.Title = "second";

        Assert.Equal(2, _document.ContentControls.Count);
        Assert.Equal(first, _document.ContentControls[0]);
        Assert.Equal(second, _document.ContentControls[1]);
        Assert.Equal(1, _document.SelectContentControlsByTag("a").Count);
        Assert.Equal(1, _document.SelectContentControlsByTitle("second").Count);
        Assert.Equal(0, _document.SelectContentControlsByTag("missing").Count);
        Assert.Equal(1, first.Range.ContentControls.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.ContentControls[2]);
    }

    [Fact]
    public void AControlInsideAnotherHasAParent()
    {
        ContentControlWrapper outer = _document.ContentControls.Add(ContentControlType.RichText, _document.Range(0, 11));
        RangeWrapper inside = outer.Range;

        Assert.Null(outer.ParentContentControl);
        Assert.Equal(outer, inside.ParentContentControl);
    }

    [Fact]
    public void DeletingAControlKeepsOrDropsItsContents()
    {
        ContentControlWrapper keep = _document.ContentControls.Add(ContentControlType.Text, _document.Range(6, 11));
        keep.Delete();

        Assert.Equal(0, _document.ContentControls.Count);
        Assert.Equal("Hello world again", _document.Content.Text.TrimEnd('\r'));
        Assert.Null(_document.WordprocessingDocument.MainDocumentPart.GlossaryDocumentPart);

        ContentControlWrapper drop = _document.ContentControls.Add(ContentControlType.Text, _document.Range(6, 11));
        drop.Delete(true);

        Assert.Equal(0, _document.ContentControls.Count);
        Assert.Equal("Hello  again", _document.Content.Text.TrimEnd('\r'));
        AssertValid();
    }

    [Fact]
    public void AControlsPlaceholderStaysWhileAnyControlUsesIt()
    {
        ContentControlWrapper one = _document.ContentControls.Add(ContentControlType.Text, _document.Range(0, 5));
        _document.ContentControls.Add(ContentControlType.Text, _document.Range(one.Range.End + 7, one.Range.End + 11));

        one.Delete();

        Assert.NotNull(_document.WordprocessingDocument.MainDocumentPart.GlossaryDocumentPart);
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.GlossaryDocumentPart.GlossaryDocument.DocParts.Elements<DocumentFormat.OpenXml.Wordprocessing.DocPart>());
    }

    [Fact]
    public void KindsTheLibraryCannotMakeAreRefused()
    {
        Assert.Throws<NotSupportedException>(() => _document.ContentControls.Add(ContentControlType.Picture, _document.Range(0, 5)));
        Assert.Throws<NotSupportedException>(() => _document.ContentControls.Add(ContentControlType.Group, _document.Range(0, 5)));
        Assert.Throws<ArgumentNullException>(() => _document.ContentControls.Add(ContentControlType.Text, null));
    }
}
