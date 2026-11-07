using System;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// InsertCells, InsertFormula and HeaderFooter of the selection. The expected grids and positions are those Word 16 gave for the same calls
/// (the selection-insert-cells-behaviour, selection-insert-formula-behaviour and selection-header-footer scenarios of
/// tools/WordReference/scenarios.ps1, which print them).
/// </summary>
public class SelectionTableMembersTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public SelectionTableMembersTest()
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

    private TableWrapper Fresh()
    {
        if (_document.Tables.Count > 0)
        {
            _document.Tables[0].Delete();
        }
        else
        {
            _document.Content.InsertAfter("Before\r");
        }
        TableWrapper table = _document.Tables.Add(_document.Range(_document.Content.End - 1, _document.Content.End - 1), 3, 3);
        for (int r = 0; r < 3; r++)
        {
            for (int c = 0; c < 3; c++)
            {
                table.Cell(r, c).Range.Text = "r" + (r + 1) + "c" + (c + 1);
            }
        }
        return table;
    }

    private string Grid()
    {
        return string.Join(" / ", _document.Tables[0].Rows.Select(r => string.Join(",", r.Cells.Select(c => c.Range.Text.TrimEnd('\r', '\a')))));
    }

    [Fact]
    public void InsertCellsMatchesWord()
    {
        TableWrapper table = Fresh();
        table.Cell(1, 1).Select();
        Sel.InsertCells(InsertCellsShift.Right);
        Assert.Equal("r1c1,r1c2,r1c3 / r2c1,,r2c2,r2c3 / r3c1,r3c2,r3c3", Grid());
        Assert.Equal((28, 29), (Sel.Start, Sel.End));

        table = Fresh();
        table.Cell(1, 1).Select();
        Sel.InsertCells(InsertCellsShift.Down);
        Assert.Equal("r1c1,r1c2,r1c3 / r2c1,,r2c3 / r3c1,r2c2,r3c3 / ,r3c2,", Grid());
        Assert.Equal((28, 29), (Sel.Start, Sel.End));

        table = Fresh();
        table.Cell(1, 1).Select();
        Sel.InsertCells(InsertCellsShift.EntireRow);
        Assert.Equal("r1c1,r1c2,r1c3 / ,, / r2c1,r2c2,r2c3 / r3c1,r3c2,r3c3", Grid());

        table = Fresh();
        table.Cell(1, 1).Select();
        Sel.InsertCells(InsertCellsShift.EntireColumn);
        Assert.Equal("r1c1,,r1c2,r1c3 / r2c1,,r2c2,r2c3 / r3c1,,r3c2,r3c3", Grid());
    }

    [Fact]
    public void InsertCellsOfSeveralCells()
    {
        TableWrapper table = Fresh();
        Sel.SetRange(table.Cell(0, 1).Range.Start, table.Cell(1, 1).Range.End - 1);
        Sel.InsertCells(InsertCellsShift.Right);
        Assert.Equal("r1c1,,r1c2,r1c3 / r2c1,,r2c2,r2c3 / r3c1,r3c2,r3c3", Grid());

        table = Fresh();
        Sel.SetRange(table.Cell(0, 1).Range.Start, table.Cell(0, 2).Range.End - 1);
        Sel.InsertCells(InsertCellsShift.Down);
        Assert.Equal("r1c1,, / r2c1,r1c2,r1c3 / r3c1,r2c2,r2c3 / ,r3c2,r3c3", Grid());

        table = Fresh();
        table.Cell(2, 2).Select();
        Sel.InsertCells(InsertCellsShift.Down);
        Assert.Equal("r1c1,r1c2,r1c3 / r2c1,r2c2,r2c3 / r3c1,r3c2, / ,,r3c3", Grid());

        table = Fresh();
        Sel.SetRange(table.Cell(1, 0).Range.Start, table.Cell(2, 1).Range.End - 1);
        Sel.InsertCells(InsertCellsShift.EntireRow);
        Assert.Equal("r1c1,r1c2,r1c3 / ,, / ,, / r2c1,r2c2,r2c3 / r3c1,r3c2,r3c3", Grid());

        table = Fresh();
        Sel.SetRange(table.Cell(1, 0).Range.Start, table.Cell(2, 1).Range.End - 1);
        Sel.InsertCells(InsertCellsShift.EntireColumn);
        Assert.Equal(",,r1c1,r1c2,r1c3 / ,,r2c1,r2c2,r2c3 / ,,r3c1,r3c2,r3c3", Grid());
    }

    [Fact]
    public void InsertCellsNeedsATable()
    {
        _document.Content.InsertAfter("Text");
        Sel.SetRange(2, 2);
        Assert.Throws<InvalidOperationException>(() => Sel.InsertCells());
        Assert.Throws<ArgumentOutOfRangeException>(() => Sel.InsertCells((InsertCellsShift)9));
    }

    [Fact]
    public void InsertFormulaReplacesTheSelection()
    {
        _document.Content.InsertAfter("One two three");
        Sel.SetRange(4, 7);
        Sel.InsertFormula();
        //Without a formula outside a cell the field has none, so that it shows an error.
        Assert.Single(_document.Fields);
        Assert.Equal("=", _document.Fields[0].Code.Text.Trim());
        Assert.StartsWith("One ", _document.Content.Text);
        Assert.EndsWith(" three\r", _document.Content.Text);
        Assert.Equal(Sel.Start, Sel.End);
        Assert.Equal(_document.Fields[0].Range.End, Sel.End);
    }

    [Fact]
    public void HeaderFooterIsTheStoryTheSelectionIsIn()
    {
        _document.Content.InsertAfter("Body text");
        Assert.Null(Sel.HeaderFooter);
        _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Head";
        _document.Sections[0].Footers[HeaderFooterIndex.Primary].Range.Text = "Foot";

        _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Select();
        Assert.True(Sel.HeaderFooter.IsHeader);
        Assert.Equal("Head", Sel.HeaderFooter.Range.Text.TrimEnd('\r'));
        _document.Sections[0].Footers[HeaderFooterIndex.Primary].Range.Select();
        Assert.False(Sel.HeaderFooter.IsHeader);
        Assert.Equal(StoryType.PrimaryFooter, Sel.StoryType);
    }
}
