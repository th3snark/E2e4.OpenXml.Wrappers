using System;
using System.Collections.Generic;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The application and its documents. The names, the order and the active document are those Word 16 gave for the same calls.
/// </summary>
public class ApplicationWrapperTest : IDisposable
{
    private readonly ApplicationWrapper _application = new ApplicationWrapper();
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");

    public void Dispose()
    {
        _application.Quit(SaveOptions.DoNotSaveChanges);
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void NewDocumentsAreNamedInOrderAndTheNewestIsActive()
    {
        Assert.Equal(0, _application.Documents.Count);
        Assert.Throws<InvalidOperationException>(() => _application.ActiveDocument);

        WordprocessingDocumentWrapper first = _application.Documents.Add();
        WordprocessingDocumentWrapper second = _application.Documents.Add();

        Assert.Equal("Document1", first.Name);
        Assert.Equal("Document2", second.Name);
        Assert.Equal(2, _application.Documents.Count);
        Assert.Same(second, _application.ActiveDocument);
        Assert.Same(second, _application.Documents[0]);
        Assert.Same(first, _application.Documents[1]);
        Assert.Same(first, _application.Documents["document1"]);
        Assert.Equal(new[] { second, first }, _application.Documents.ToArray());
        Assert.Same(_application, first.Application);
        Assert.Equal("Microsoft Word", _application.Name);
    }

    [Fact]
    public void ActivatingChangesTheActiveDocumentButNotTheOrder()
    {
        WordprocessingDocumentWrapper first = _application.Documents.Add();
        WordprocessingDocumentWrapper second = _application.Documents.Add();

        first.Activate();

        Assert.Same(first, _application.ActiveDocument);
        Assert.Same(second, _application.Documents[0]);
        Assert.Same(first.Selection, _application.Selection);
    }

    [Fact]
    public void ANameIsNotUsedAgainAfterItsDocumentIsClosed()
    {
        WordprocessingDocumentWrapper first = _application.Documents.Add();
        WordprocessingDocumentWrapper second = _application.Documents.Add();
        WordprocessingDocumentWrapper third = _application.Documents.Add();

        third.Close(SaveOptions.DoNotSaveChanges);
        Assert.Equal(2, _application.Documents.Count);
        Assert.Same(second, _application.ActiveDocument);

        Assert.Equal("Document4", _application.Documents.Add().Name);
        Assert.Throws<KeyNotFoundException>(() => _application.Documents["Document3"]);
        first.Close(SaveOptions.DoNotSaveChanges);
        Assert.Equal(2, _application.Documents.Count);
    }

    [Fact]
    public void ClosingTheActiveDocumentActivatesAnother()
    {
        WordprocessingDocumentWrapper first = _application.Documents.Add();
        WordprocessingDocumentWrapper second = _application.Documents.Add();

        second.Close(SaveOptions.DoNotSaveChanges);
        Assert.Same(first, _application.ActiveDocument);

        first.Close(SaveOptions.DoNotSaveChanges);
        Assert.Equal(0, _application.Documents.Count);
        Assert.Throws<InvalidOperationException>(() => _application.ActiveDocument);
    }

    [Fact]
    public void ADocumentIsOpenedOnceAndSavedBack()
    {
        using (WordprocessingDocumentWrapper created = WordprocessingDocumentWrapper.Create(_path, true, false))
        {
            created.Content.InsertAfter("Hello");
        }

        WordprocessingDocumentWrapper opened = _application.Documents.Open(_path);
        Assert.Equal(System.IO.Path.GetFileName(_path), opened.Name);
        Assert.Same(opened, _application.Documents.Open(_path));
        Assert.Equal(1, _application.Documents.Count);
        Assert.False(opened.ReadOnly);
        opened.Content.InsertAfter(" world");

        _application.Documents.Save();
        _application.Documents.Close();

        Assert.Equal(0, _application.Documents.Count);
        using (WordprocessingDocumentWrapper again = WordprocessingDocumentWrapper.Open(_path))
        {
            Assert.Equal("Hello world\r", again.Content.Text);
        }
    }

    [Fact]
    public void ADocumentOpenedReadOnlyIsNotSaved()
    {
        using (WordprocessingDocumentWrapper created = WordprocessingDocumentWrapper.Create(_path, true, false))
        {
            created.Content.InsertAfter("Hello");
        }

        WordprocessingDocumentWrapper opened = _application.Documents.Open(_path, true);

        Assert.True(opened.ReadOnly);
        _application.Documents.Save();
        _application.Quit();
        Assert.Equal(0, _application.Documents.Count);
    }

    [Fact]
    public void TheNameAndInitialsOfTheUserAreThoseOfTheCommentsAndDoNotFollowEachOther()
    {
        WordprocessingDocumentWrapper document = _application.Documents.Add();
        document.Content.InsertAfter("Hello world");

        _application.UserName = "Probe Person";
        _application.UserInitials = "ZZ";
        CommentWrapper comment = document.Comments.Add(document.Range(0, 5), "Note");
        Assert.Equal("Probe Person", comment.Author);
        Assert.Equal("ZZ", comment.Initial);

        _application.UserName = "Other Name";
        Assert.Equal("ZZ", _application.UserInitials);
    }

    [Fact]
    public void UnitsAreConvertedAsWordDoes()
    {
        Assert.Equal(28.34646, _application.CentimetersToPoints(1), 4);
        Assert.Equal(70.86614, _application.CentimetersToPoints(2.5), 4);
        Assert.Equal(72, _application.InchesToPoints(1), 4);
        Assert.Equal(108, _application.InchesToPoints(1.5), 4);
        Assert.Equal(12, _application.LinesToPoints(1), 4);
        Assert.Equal(2.834646, _application.MillimetersToPoints(1), 4);
        Assert.Equal(19.84252, _application.MillimetersToPoints(7), 4);
        Assert.Equal(12, _application.PicasToPoints(1), 4);
        Assert.Equal(2.54, _application.PointsToCentimeters(72), 4);
        Assert.Equal(1.388889, _application.PointsToInches(100), 4);
        Assert.Equal(2, _application.PointsToLines(24), 4);
        Assert.Equal(25.4, _application.PointsToMillimeters(72), 4);
        Assert.Equal(2.5, _application.PointsToPicas(30), 4);
        Assert.Equal("16.0", _application.Version);
        Assert.Equal(string.Empty, _application.UserAddress);
        _application.UserAddress = "1 Road";
        Assert.Equal("1 Road", _application.UserAddress);
        Assert.Throws<ArgumentNullException>(() => _application.UserAddress = null);
    }

    [Fact]
    public void TheUserStartsAsTheUserOfTheSystem()
    {
        Assert.False(string.IsNullOrEmpty(_application.UserName));
        Assert.False(string.IsNullOrEmpty(_application.UserInitials));
    }

    [Fact]
    public void ADocumentMadeOutsideAnApplicationHasOneOfItsOwn()
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(_path, true, false))
        {
            ApplicationWrapper own = document.Application;
            Assert.Same(own, document.Application);
            Assert.Same(document, own.ActiveDocument);
            Assert.Equal(1, own.Documents.Count);
            own.UserName = "Own Name";
            document.Content.InsertAfter("x");
            Assert.Equal("Own Name", document.Comments.Add(document.Range(0, 1), "c").Author);
        }
    }

    [Fact]
    public void QuitClosesTheDocumentsAndTheArgumentsAreChecked()
    {
        WordprocessingDocumentWrapper document = _application.Documents.Add();

        Assert.Throws<NotSupportedException>(() => _application.Quit(SaveOptions.PromptToSaveChanges));
        Assert.Throws<NotSupportedException>(() => _application.Documents.Close(SaveOptions.PromptToSaveChanges));
        Assert.Throws<ArgumentOutOfRangeException>(() => _application.Quit((SaveOptions)99));
        Assert.Throws<ArgumentOutOfRangeException>(() => _application.Documents[1]);
        Assert.Throws<ArgumentNullException>(() => _application.Documents[null]);
        Assert.Throws<ArgumentNullException>(() => _application.Documents.Open(null));
        Assert.Throws<ArgumentNullException>(() => _application.UserName = null);

        _application.Quit(SaveOptions.DoNotSaveChanges);
        Assert.Equal(0, _application.Documents.Count);
        Assert.NotNull(document);
    }
}
