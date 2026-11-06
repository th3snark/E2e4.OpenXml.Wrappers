using System;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The selection. The expected positions and texts are those Word 16 gave for the same calls (the selection-* scenarios of
/// tools/WordReference/scenarios.ps1, which print them).
/// </summary>
public class SelectionWrapperTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public SelectionWrapperTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private SelectionWrapper Sel
    {
        get { return _document.Selection; }
    }

    private string DocumentText
    {
        get { return _document.Content.Text; }
    }

    private void AssertAt(int start, int end, string text)
    {
        Assert.Equal(start, Sel.Start);
        Assert.Equal(end, Sel.End);
        Assert.Equal(text, Sel.Text);
    }

    [Fact]
    public void TypingBuildsTheTextAndKeepsTheInsertionPointAfterIt()
    {
        AssertAt(0, 0, "\r");
        Assert.Equal(SelectionType.InsertionPoint, Sel.Type);

        Sel.TypeText("Hello world again");
        AssertAt(17, 17, "\r");
        Assert.Equal("Hello world again\r", DocumentText);

        Sel.TypeParagraph();
        AssertAt(18, 18, "\r");
        Sel.TypeText("Second");
        AssertAt(24, 24, "\r");
        Assert.Equal("Hello world again\rSecond\r", DocumentText);

        Sel.TypeText(string.Empty);
        AssertAt(24, 24, "\r");
    }

    [Fact]
    public void KeysGoToTheStartAndTheEndOfTheStory()
    {
        Sel.TypeText("Hello world again\rSecond");

        Assert.Equal(24, Sel.HomeKey(KeyUnit.Story));
        AssertAt(0, 0, "H");
        Assert.Equal(24, Sel.EndKey(KeyUnit.Story));
        AssertAt(24, 24, "\r");

        Sel.HomeKey(KeyUnit.Story);
        Sel.EndKey(KeyUnit.Story, Movement.Extend);
        AssertAt(0, 24, "Hello world again\rSecond");
        Assert.Equal(SelectionType.Normal, Sel.Type);
    }

    [Fact]
    public void MovingByCharactersWordsAndExtendingFollowsWord()
    {
        Sel.TypeText("Hello world again\rSecond line here\rThird");
        _document.Range(6, 11).Select();
        AssertAt(6, 11, "world");

        Assert.Equal(1, Sel.MoveRight(TextUnit.Character, 1));
        AssertAt(11, 11, " ");

        _document.Range(6, 11).Select();
        Sel.MoveLeft(TextUnit.Character, 1);
        AssertAt(6, 6, "w");

        _document.Range(6, 11).Select();
        Sel.MoveRight(TextUnit.Character, 3);
        AssertAt(13, 13, "g");

        _document.Range(8, 8).Select();
        Assert.Equal(3, Sel.MoveRight(TextUnit.Character, 3, Movement.Extend));
        AssertAt(8, 11, "rld");

        Sel.MoveLeft(TextUnit.Character, 5, Movement.Extend);
        AssertAt(6, 8, "wo");

        Sel.MoveRight(TextUnit.Character, 1, Movement.Extend);
        AssertAt(7, 8, "o");

        _document.Range(8, 8).Select();
        Sel.MoveRight(TextUnit.Word, 1);
        AssertAt(12, 12, "a");
        Sel.MoveRight(TextUnit.Word, 1, Movement.Extend);
        AssertAt(12, 17, "again");
    }

    [Fact]
    public void MovingNeverLeavesTheStory()
    {
        Sel.TypeText("Hello");

        Assert.Equal(3, Sel.MoveLeft(TextUnit.Character, 3));
        Assert.Equal(2, Sel.MoveLeft(TextUnit.Character, 100));
        AssertAt(0, 0, "H");
        Sel.EndKey(KeyUnit.Story);
        Assert.Equal(0, Sel.MoveRight(TextUnit.Character, 3));
        AssertAt(5, 5, "\r");
        Assert.Throws<ArgumentException>(() => Sel.MoveRight(TextUnit.Paragraph, 1));
    }

    [Fact]
    public void ParagraphsAreMovedByUpAndDown()
    {
        Sel.TypeText("Hello world again\rSecond line here\rThird");

        _document.Range(25, 25).Select();
        Assert.Equal(1, Sel.MoveUp(TextUnit.Paragraph, 1));
        AssertAt(18, 18, "S");
        Sel.MoveUp(TextUnit.Paragraph, 1);
        AssertAt(0, 0, "H");

        Sel.MoveDown(TextUnit.Paragraph, 2);
        AssertAt(35, 35, "T");

        _document.Range(2, 2).Select();
        Sel.MoveDown(TextUnit.Paragraph, 1, Movement.Extend);
        AssertAt(2, 18, "llo world again\r");

        Assert.Throws<NotSupportedException>(() => Sel.MoveDown(TextUnit.Sentence, 1));
    }

    [Fact]
    public void TypingOverAndDeletingAreDoneOnTheSelection()
    {
        Sel.TypeText("Hello world again");
        _document.Range(6, 11).Select();

        Sel.TypeText("X");
        AssertAt(7, 7, " ");
        Assert.Equal("Hello X again\r", DocumentText);

        Sel.TypeBackspace();
        AssertAt(6, 6, " ");
        Assert.Equal("Hello  again\r", DocumentText);

        _document.Range(0, 5).Select();
        Sel.TypeBackspace();
        AssertAt(0, 0, " ");
        Assert.Equal("  again\r", DocumentText);
        Sel.TypeBackspace();
        AssertAt(0, 0, " ");
    }

    [Fact]
    public void DeleteTakesTheCharacterAfterTheInsertionPointOrTheSelection()
    {
        Sel.TypeText("Hello world again");
        Sel.HomeKey(KeyUnit.Story);
        Sel.MoveRight(TextUnit.Character, 3);

        Assert.Equal(1, Sel.Delete());
        AssertAt(3, 3, "o");
        Assert.Equal("Helo world again\r", DocumentText);

        Assert.Equal(1, Sel.Delete(TextUnit.Word, 1));

        _document.Range(0, 3).Select();
        Assert.Equal(1, Sel.Delete());
        Assert.Equal(Sel.Start, Sel.End);
        Sel.EndKey(KeyUnit.Story);
        Assert.Equal(0, Sel.Delete());
    }

    [Fact]
    public void InsertingExtendsTheSelectionOverTheNewText()
    {
        Sel.TypeText("Hello world again");
        Sel.HomeKey(KeyUnit.Story);
        Sel.MoveRight(TextUnit.Character, 3);

        Sel.InsertAfter("Z");
        AssertAt(3, 4, "Z");
        Assert.Equal(SelectionType.Normal, Sel.Type);

        Sel.InsertBefore("Y");
        AssertAt(3, 5, "YZ");
    }

    [Fact]
    public void TheSelectionFollowsEditsMadeElsewhere()
    {
        Sel.TypeText("Hello world again");
        _document.Range(0, 3).Select();

        _document.Range(2, 2).InsertAfter("QQ");

        AssertAt(0, 5, "HeQQl");
        Sel.SetRange(1, 4);
        AssertAt(1, 4, "eQQ");
        Assert.Equal("eQQ", Sel.Range.Text);
        Assert.NotSame(Sel.Range, Sel.Range);
    }

    [Fact]
    public void ExpandingAndCollapsingWork()
    {
        Sel.TypeText("Hello world again\rS");
        Sel.EndKey(KeyUnit.Story);
        AssertAt(19, 19, "\r");

        Sel.Expand(TextUnit.Word);
        AssertAt(19, 20, "\r");
        Sel.Expand(TextUnit.Paragraph);
        AssertAt(18, 20, "S\r");
        Sel.Collapse(CollapseDirection.End);
        Assert.Equal(Sel.Start, Sel.End);
        Sel.Collapse(CollapseDirection.Start);
        Assert.Equal(SelectionType.InsertionPoint, Sel.Type);
    }

    [Fact]
    public void TablesAreSelectedByCellRowAndKeys()
    {
        _document.Content.InsertAfter("Before\r");
        _document.Tables.Add(_document.Range(_document.Content.End - 1, _document.Content.End - 1), 2, 2);
        _document.Tables[0].Cell(0, 0).Range.Text = "a";
        _document.Tables[0].Cell(0, 1).Range.Text = "b";
        _document.Tables[0].Cell(1, 0).Range.Text = "c";
        _document.Tables[0].Cell(1, 1).Range.Text = "d";

        _document.Tables[0].Cell(0, 1).Select();
        Assert.Equal(9, Sel.Start);
        Assert.Equal(11, Sel.End);
        Assert.Equal(SelectionType.Column, Sel.Type);

        Sel.SelectRow();
        Assert.Equal(7, Sel.Start);
        Assert.Equal(12, Sel.End);
        Assert.Equal(SelectionType.Row, Sel.Type);

        Sel.Collapse(CollapseDirection.Start);
        Sel.MoveRight(TextUnit.Character, 2);
        Sel.HomeKey(KeyUnit.Row);
        Assert.Equal(7, Sel.Start);
        Sel.EndKey(KeyUnit.Row);
        Assert.Equal(9, Sel.Start);
        Sel.EndKey(KeyUnit.Column);
        Assert.Equal(14, Sel.Start);
        Sel.HomeKey(KeyUnit.Column);
        Assert.Equal(9, Sel.Start);

        Sel.SelectCell();
        Assert.Equal(9, Sel.Start);
        Assert.Equal(11, Sel.End);

        _document.Range(0, 0).Select();
        Assert.Throws<InvalidOperationException>(() => Sel.SelectRow());
        Assert.Throws<InvalidOperationException>(() => Sel.EndKey(KeyUnit.Row));
    }

    [Fact]
    public void ThereIsOneSelectionForEachDocument()
    {
        Assert.Same(_document.Selection, _document.Selection);
        Assert.Same(_document, _document.Selection.Document);
        Assert.Throws<ArgumentNullException>(() => _document.Selection.TypeText(null));
    }
}
