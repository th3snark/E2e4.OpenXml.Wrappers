using System;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Extending the selection, moving to fields, and the table members of the selection. The expected positions and texts are those Word 16
/// gave for the same calls (the selection-extend-units, selection-extend-mode, selection-extend-case, selection-next-field,
/// selection-insert-columns-behaviour and selection-insert-rows-behaviour scenarios of tools/WordReference/scenarios.ps1, which print them).
/// </summary>
public class SelectionExtendTest : IDisposable
{
    private const string Text = "One two three. Four five six.\rSecond paragraph here.\rThird.";

    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public SelectionExtendTest()
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

    private void AssertAt(int start, int end, string text)
    {
        Assert.Equal(start, Sel.Start);
        Assert.Equal(end, Sel.End);
        Assert.Equal(text, Sel.Text);
    }

    [Fact]
    public void ExtendGrowsTheSelectionByUnitsAndShrinkTakesItBack()
    {
        _document.Content.InsertAfter(Text);
        Sel.SetRange(5, 5);
        Assert.False(Sel.ExtendMode);

        Sel.Extend();
        Assert.True(Sel.ExtendMode);
        AssertAt(5, 5, "w");
        Sel.Extend();
        AssertAt(4, 8, "two ");
        Sel.Extend();
        AssertAt(0, 15, "One two three. ");
        Sel.Extend();
        AssertAt(0, 30, "One two three. Four five six.\r");
        Sel.Extend();
        AssertAt(0, 60, Text + "\r");
        Sel.Extend();
        AssertAt(0, 60, Text + "\r");

        Sel.Shrink();
        AssertAt(0, 30, "One two three. Four five six.\r");
        Sel.Shrink();
        AssertAt(0, 15, "One two three. ");
        Sel.Shrink();
        AssertAt(4, 8, "two ");
        Sel.Shrink();
        AssertAt(5, 5, "w");
        Sel.Shrink();
        AssertAt(5, 5, "w");
        Assert.True(Sel.ExtendMode);
    }

    [Fact]
    public void ExtendToACharacterIncludesItAndKeepsTheStartWhileItContinues()
    {
        _document.Content.InsertAfter(Text);
        Sel.SetRange(5, 5);

        Sel.Extend("e");
        AssertAt(5, 12, "wo thre");
        Sel.Extend("e");
        AssertAt(5, 13, "wo three");
        Sel.Extend(".");
        AssertAt(5, 14, "wo three.");
        Sel.Extend("\r");
        AssertAt(5, 30, "wo three. Four five six.\r");
        Sel.Extend("Q");
        AssertAt(5, 30, "wo three. Four five six.\r");
        Assert.False(Sel.ExtendMode);
    }

    [Fact]
    public void ExtendToACharacterCountsTheCaseAndTakesOneCharacter()
    {
        _document.Content.InsertAfter("one Two three Four");
        Sel.SetRange(0, 0);
        Sel.Extend("t");
        Assert.Equal((0, 9), (Sel.Start, Sel.End));
        Sel.SetRange(0, 0);
        Sel.Extend("T");
        Assert.Equal((0, 5), (Sel.Start, Sel.End));
        Sel.SetRange(0, 0);
        Sel.Extend("tw");
        Assert.Equal((0, 0), (Sel.Start, Sel.End));

        //A selection that Extend did not make is dropped: the new one starts where it ended.
        Sel.SetRange(0, 3);
        Sel.Extend("o");
        Assert.Equal((3, 7), (Sel.Start, Sel.End));
    }

    [Fact]
    public void ShrinkOfASelectionNotMadeByExtendUsesItsStart()
    {
        _document.Content.InsertAfter(Text);
        Sel.SetRange(5, 30);
        Sel.Shrink();
        AssertAt(0, 15, "One two three. ");

        Sel.SetRange(5, 8);
        Sel.Shrink();
        AssertAt(5, 5, "w");
    }

    [Fact]
    public void MovingExtendsWhileExtendModeIsOnAndTypingTurnsItOff()
    {
        _document.Content.InsertAfter(Text);
        Sel.SetRange(5, 5);
        Sel.ExtendMode = true;

        Sel.MoveRight(TextUnit.Character, 3);
        AssertAt(5, 8, "wo ");
        Sel.MoveLeft(TextUnit.Character, 1);
        AssertAt(5, 7, "wo");
        Sel.MoveRight(TextUnit.Word, 1);
        AssertAt(5, 8, "wo ");
        Sel.MoveDown(TextUnit.Paragraph, 1);
        AssertAt(5, 30, "wo three. Four five six.\r");

        //Collapsing and setting the range leave the mode on.
        Sel.Collapse(CollapseDirection.Start);
        AssertAt(5, 5, "w");
        Assert.True(Sel.ExtendMode);
        Sel.SetRange(2, 6);
        Assert.True(Sel.ExtendMode);

        Sel.Collapse(CollapseDirection.Start);
        Sel.TypeText("X");
        Assert.False(Sel.ExtendMode);
    }

    [Fact]
    public void ExtendFromAnInsertionPointWithTheModeOnSelectsTheWord()
    {
        _document.Content.InsertAfter(Text);
        Sel.SetRange(16, 16);
        Sel.ExtendMode = true;
        Sel.Extend();
        AssertAt(15, 20, "Four ");
    }

    private (FieldWrapper First, FieldWrapper Second) TwoFields()
    {
        _document.Content.InsertAfter("A ");
        FieldWrapper first = _document.Fields.Add(_document.Range(_document.Content.End - 1, _document.Content.End - 1), FieldType.Empty, "DATE", false);
        _document.Content.InsertAfter(" B ");
        FieldWrapper second = _document.Fields.Add(_document.Range(_document.Content.End - 1, _document.Content.End - 1), FieldType.Empty, "PAGE", false);
        _document.Content.InsertAfter(" C");
        return (first, second);
    }

    private void AssertSelects(FieldWrapper field)
    {
        Assert.NotNull(field);
        Assert.Equal((field.Range.Start, field.Range.End), (Sel.Start, Sel.End));
    }

    [Fact]
    public void NextFieldAndPreviousFieldSelectTheNeighbouringField()
    {
        (FieldWrapper first, FieldWrapper second) = TwoFields();
        int firstStart = first.Range.Start;
        int firstEnd = first.Range.End;
        int secondStart = second.Range.Start;
        int secondEnd = second.Range.End;

        Sel.SetRange(0, 0);
        AssertSelects(Sel.NextField());
        Assert.Equal((firstStart, firstEnd), (Sel.Start, Sel.End));
        AssertSelects(Sel.NextField());
        Assert.Equal((secondStart, secondEnd), (Sel.Start, Sel.End));

        //Nothing after the last field: nothing is found and the selection stays.
        Assert.Null(Sel.NextField());
        Assert.Equal((secondStart, secondEnd), (Sel.Start, Sel.End));

        AssertSelects(Sel.PreviousField());
        Assert.Equal((firstStart, firstEnd), (Sel.Start, Sel.End));
        Assert.Null(Sel.PreviousField());
        Assert.Equal((firstStart, firstEnd), (Sel.Start, Sel.End));
    }

    [Fact]
    public void NextFieldSkipsTheFieldTheInsertionPointIsInsideAndPreviousFieldNeedsAFieldBeforeIt()
    {
        (FieldWrapper first, FieldWrapper second) = TwoFields();
        int firstStart = first.Range.Start;
        int firstEnd = first.Range.End;
        int secondStart = second.Range.Start;
        int secondEnd = second.Range.End;

        Sel.SetRange(firstStart + 1, firstStart + 1);
        AssertSelects(Sel.NextField());
        Assert.Equal((secondStart, secondEnd), (Sel.Start, Sel.End));

        Sel.SetRange(_document.Content.End - 1, _document.Content.End - 1);
        AssertSelects(Sel.PreviousField());
        Assert.Equal((secondStart, secondEnd), (Sel.Start, Sel.End));

        Sel.SetRange(firstStart, firstStart);
        AssertSelects(Sel.NextField());
        Assert.Equal((firstStart, firstEnd), (Sel.Start, Sel.End));

        Sel.SetRange(firstStart + 1, firstStart + 1);
        Assert.Null(Sel.PreviousField());
        Assert.Equal((firstStart + 1, firstStart + 1), (Sel.Start, Sel.End));

        Sel.SetRange(firstEnd, firstEnd);
        AssertSelects(Sel.NextField());
        Assert.Equal((secondStart, secondEnd), (Sel.Start, Sel.End));

        Sel.SetRange(0, firstStart + 3);
        AssertSelects(Sel.NextField());
        Assert.Equal((firstStart, firstEnd), (Sel.Start, Sel.End));
    }

    private TableWrapper ThreeColumns()
    {
        _document.Content.InsertAfter("Before\r");
        TableWrapper table = _document.Tables.Add(_document.Range(_document.Content.End - 1, _document.Content.End - 1), 2, 3);
        for (int r = 0; r < 2; r++)
        {
            for (int c = 0; c < 3; c++)
            {
                table.Cell(r, c).Range.Text = "r" + (r + 1) + "c" + (c + 1);
            }
        }
        return table;
    }

    private string FirstRow()
    {
        return string.Join("|", _document.Tables[0].Rows[0].Cells.Select(c => c.Range.Text.TrimEnd('\r', '\a')));
    }

    [Fact]
    public void InsertingColumnsSelectsTheNewColumns()
    {
        TableWrapper table = ThreeColumns();
        Assert.Equal("r1c1|r1c2|r1c3", FirstRow());

        table.Cell(0, 1).Select();
        Sel.InsertColumns();
        Assert.Equal("r1c1||r1c2|r1c3", FirstRow());
        Assert.Equal((12, 30, SelectionType.Column), (Sel.Start, Sel.End, Sel.Type));

        _document.Tables[0].Cell(0, 1).Select();
        Sel.InsertColumnsRight();
        Assert.Equal("r1c1|||r1c2|r1c3", FirstRow());
        Assert.Equal((13, 32, SelectionType.Column), (Sel.Start, Sel.End, Sel.Type));

        //Two selected cells make two columns.
        Sel.SetRange(_document.Tables[0].Cell(0, 0).Range.Start, _document.Tables[0].Cell(0, 1).Range.End);
        Assert.Equal((7, 13, SelectionType.Column), (Sel.Start, Sel.End, Sel.Type));
        Sel.InsertColumns();
        Assert.Equal("||r1c1|||r1c2|r1c3", FirstRow());
        Assert.Equal((7, 29, SelectionType.Column), (Sel.Start, Sel.End, Sel.Type));
        Assert.Equal(7, _document.Tables[0].Columns.Count);

        _document.Tables[0].Cell(0, 2).Select();
        Sel.SelectColumn();
        Assert.Equal((9, 34, SelectionType.Column), (Sel.Start, Sel.End, Sel.Type));
        Sel.InsertColumnsRight();
        Assert.Equal("||r1c1||||r1c2|r1c3", FirstRow());
        Assert.Equal((14, 36, SelectionType.Column), (Sel.Start, Sel.End, Sel.Type));
        Assert.Equal(8, _document.Tables[0].Columns.Count);
    }

    //Not measured in Word: a merged cell that spans the place of the new column grows, so that no text is lost.
    [Fact]
    public void AMergedCellThatSpansTheNewColumnGrows()
    {
        TableWrapper table = ThreeColumns();
        table.Cell(1, 0).Merge(table.Cell(1, 1));
        table.Cell(0, 1).Select();
        Sel.InsertColumns();

        TableWrapper after = _document.Tables[0];
        Assert.Equal(4, after.Columns.Count);
        Assert.Equal(4, after.Rows[0].Cells.Count);
        Assert.Equal(2, after.Rows[1].Cells.Count);
        Assert.Equal("r2c1\rr2c2\a".Replace("\a", string.Empty), after.Rows[1].Cells[0].Range.Text.TrimEnd('\r', '\a').Replace("\r\r", "\r"));
    }
    [Fact]
    public void InsertingRowsSelectsTheNewRows()
    {
        _document.Content.InsertAfter("Before\r");
        TableWrapper table = _document.Tables.Add(_document.Range(_document.Content.End - 1, _document.Content.End - 1), 3, 2);
        for (int r = 0; r < 3; r++)
        {
            for (int c = 0; c < 2; c++)
            {
                table.Cell(r, c).Range.Text = "r" + (r + 1) + "c" + (c + 1);
            }
        }
        table.Cell(1, 0).Select();
        Sel.InsertRows(2);
        Assert.Equal(5, _document.Tables[0].Rows.Count);
        Assert.Equal("r1c1|||r2c1|r3c1", string.Join("|", _document.Tables[0].Columns[0].Cells.Select(c => c.Range.Text.TrimEnd('\r', '\a'))));
        Assert.Equal((18, 24, SelectionType.Row), (Sel.Start, Sel.End, Sel.Type));

        _document.Tables[0].Cell(0, 0).Select();
        Sel.InsertRowsAbove(1);
        Assert.Equal((7, 10, SelectionType.Row), (Sel.Start, Sel.End, Sel.Type));
        _document.Tables[0].Cell(0, 0).Select();
        Sel.InsertRowsBelow(1);
        Assert.Equal((10, 13, SelectionType.Row), (Sel.Start, Sel.End, Sel.Type));
    }

    [Fact]
    public void TableMembersNeedATable()
    {
        _document.Content.InsertAfter("Text");
        Sel.SetRange(2, 2);
        Assert.Throws<InvalidOperationException>(() => Sel.InsertRows(2));
        Assert.Throws<InvalidOperationException>(() => Sel.InsertColumns());
        Assert.Throws<InvalidOperationException>(() => Sel.InsertColumnsRight());
        Assert.Throws<InvalidOperationException>(() => Sel.SelectColumn());
        Assert.Throws<ArgumentOutOfRangeException>(() => Sel.InsertRows(0));
    }
}
