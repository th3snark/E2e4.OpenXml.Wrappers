using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class SelectionWrapper
{
    #region Clearing formatting

    /// <summary>
    /// Takes the direct character formatting off the selection so that its styles apply, as <c>Selection.ClearCharacterDirectFormatting</c>
    /// in Word. A character style that has been applied is kept.
    /// </summary>
    public void ClearCharacterDirectFormatting()
    {
        Font.Reset();
    }

    /// <summary>Takes the character style off the selection, as <c>Selection.ClearCharacterStyle</c> in Word. Direct formatting is kept.</summary>
    public void ClearCharacterStyle()
    {
        Font.ClearStyle();
    }

    /// <summary>Takes the character style and the direct character formatting off the selection, as <c>Selection.ClearCharacterAllFormatting</c> in Word.</summary>
    public void ClearCharacterAllFormatting()
    {
        ClearCharacterDirectFormatting();
        ClearCharacterStyle();
    }

    /// <summary>
    /// Takes the direct paragraph formatting off the paragraphs of the selection so that their styles apply, as
    /// <c>Selection.ClearParagraphDirectFormatting</c> in Word. The paragraph style is kept.
    /// </summary>
    public void ClearParagraphDirectFormatting()
    {
        ParagraphFormat.Reset();
    }

    /// <summary>Gives the paragraphs of the selection the Normal style, as <c>Selection.ClearParagraphStyle</c> in Word. Direct formatting is kept.</summary>
    public void ClearParagraphStyle()
    {
        foreach (Paragraph paragraph in _range.GetParagraphs())
        {
            paragraph.ParagraphProperties?.RemoveAllChildren<ParagraphStyleId>();
            paragraph.RemoveEmptyProperties();
        }
    }

    /// <summary>Takes the paragraph style and the direct paragraph formatting off the paragraphs of the selection, as <c>Selection.ClearParagraphAllFormatting</c> in Word.</summary>
    public void ClearParagraphAllFormatting()
    {
        ClearParagraphDirectFormatting();
        ClearParagraphStyle();
    }

    /// <summary>
    /// Takes the character and paragraph styles and the direct formatting off the selection, as <c>Selection.ClearFormatting</c> in Word.
    /// </summary>
    public void ClearFormatting()
    {
        ClearCharacterAllFormatting();
        ClearParagraphAllFormatting();
    }

    /// <summary>
    /// Makes the selection bold, or not bold when all of it is bold already, as <c>Selection.BoldRun</c> in Word. The text is bold for complex
    /// script too.
    /// </summary>
    public void BoldRun()
    {
        bool on = Font.Bold != true;
        Font.Bold = on;
        Font.BoldBi = on;
    }

    /// <summary>
    /// Makes the selection italic, or not italic when all of it is italic already, as <c>Selection.ItalicRun</c> in Word. The text is italic
    /// for complex script too.
    /// </summary>
    public void ItalicRun()
    {
        bool on = Font.Italic != true;
        Font.Italic = on;
        Font.ItalicBi = on;
    }

    /// <summary>Makes the paragraphs of the selection read right to left, as <c>Selection.RtlPara</c> in Word.</summary>
    public void RtlPara()
    {
        ParagraphFormat.ReadingOrder = ReadingOrder.RightToLeft;
    }

    /// <summary>Makes the paragraphs of the selection read left to right, as <c>Selection.LtrPara</c> in Word.</summary>
    public void LtrPara()
    {
        ParagraphFormat.ReadingOrder = ReadingOrder.LeftToRight;
    }

    /// <summary>Marks the selected text as reading right to left, as <c>Selection.RtlRun</c> in Word.</summary>
    public void RtlRun()
    {
        SetRunDirection(true);
    }

    /// <summary>Marks the selected text as reading left to right, as <c>Selection.LtrRun</c> in Word.</summary>
    public void LtrRun()
    {
        SetRunDirection(false);
    }

    //Word marks the paragraph mark too when the selection reaches the end of the text of its paragraph.
    private void SetRunDirection(bool rightToLeft)
    {
        Font.RightToLeftRun = rightToLeft;
        if (_range.End == _range.Start)
        {
            return;
        }
        RangeWrapper last = _range.Duplicate();
        last.SetRange(_range.End, _range.End);
        ParagraphWrapper paragraph = last.Paragraphs[0];
        if (paragraph.Range.End - 1 != _range.End)
        {
            return;
        }
        ParagraphMarkRunProperties mark = paragraph.Paragraph.GetOrAddParagraphMarkRunProperties();
        mark.RemoveRunProperty("rtl");
        if (rightToLeft)
        {
            mark.SetRunProperty(new RightToLeftText());
        }
        if (!mark.HasChildren)
        {
            mark.Remove();
        }
        paragraph.Paragraph.RemoveEmptyProperties();
    }
    #endregion

    #region Rows

    //The rows of the tables that the selection touches, in order, or the row of the cell the insertion point is in.
    private List<TableRowWrapper> SelectedRows()
    {
        return SelectedCells().Select(c => c.Row).Distinct().ToList();
    }

    //Selects rows that were just added.
    private void SelectAdded(List<TableRowWrapper> rows)
    {
        SelectBetween(rows[0].Range, rows[rows.Count - 1].Range);
    }

    /// <summary>
    /// Inserts rows below the last row of the selection, as <c>Selection.InsertRowsBelow</c> in Word. The new rows have the formatting of that row.
    /// </summary>
    /// <param name="numRows">How many rows to insert.</param>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="numRows"/> is less than one.</exception>
    public void InsertRowsBelow(int numRows = 1)
    {
        Guard.ThrowIfLessThan(numRows, 1);
        TableRowWrapper last = SelectedRows().Last();
        TableRowWrapper next = last.Next();
        List<TableRowWrapper> added = new List<TableRowWrapper>();
        for (int i = 0; i < numRows; i++)
        {
            added.Add(last.Parent.Rows.Add(next));
        }
        SelectAdded(added);
    }

    /// <summary>
    /// Inserts rows above the first row of the selection, as <c>Selection.InsertRowsAbove</c> in Word. The new rows have the formatting of that row.
    /// </summary>
    /// <param name="numRows">How many rows to insert.</param>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="numRows"/> is less than one.</exception>
    public void InsertRowsAbove(int numRows = 1)
    {
        Guard.ThrowIfLessThan(numRows, 1);
        TableRowWrapper first = SelectedRows().First();
        List<TableRowWrapper> added = new List<TableRowWrapper>();
        for (int i = 0; i < numRows; i++)
        {
            added.Add(first.Parent.Rows.Add(first));
        }
        SelectAdded(added);
    }

    #endregion
}
