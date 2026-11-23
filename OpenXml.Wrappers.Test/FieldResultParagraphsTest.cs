using System.IO;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// A field's result can be several paragraphs, as in Word (#97): setting <c>Field.Result.Text</c> with paragraph marks splits the result into
/// paragraphs, in the formatting of the field's own paragraph; a paragraph mark at the end puts the field's end in a paragraph of its own; and
/// setting text without paragraph marks joins them again. What Word 16 gave is in the comments.
/// </summary>
public class FieldResultParagraphsTest
{
    private static WordprocessingDocumentWrapper Create(Stream stream, out FieldWrapper field, int at = -1)
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Content.Text = "Before\rAfter";
        int position = at < 0 ? document.Paragraphs[0].Range.End - 1 : at;
        field = document.Fields.Add(document.Range(position, position), FieldType.Empty, "QUOTE \"x\"", false);
        Assert.True(field.Update());
        return document;
    }

    //Word: the result "A\rB\rC" is 19 to 24, three paragraphs; the document "BeforeA\rB\rC\rAfter\r".
    [Fact]
    public void ResultWithParagraphMarksIsSeveralParagraphs()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field);

        field.Result.Text = "A\rB\rC";

        Assert.Equal("A\rB\rC", field.Result.Text);
        Assert.Equal(19, field.Result.Range.Start);
        Assert.Equal(24, field.Result.Range.End);
        Assert.Equal(3, field.Result.Range.Paragraphs.Count);
        Assert.Equal("BeforeA\rB\rC\rAfter\r", document.Content.Text);
        Assert.Equal(1, document.Fields.Count);
        Assert.Equal("QUOTE \"x\"", field.Code.Text);
    }

    //Word: "A\rB\rC\r" puts the end of the field in a paragraph of its own; the result is 19 to 25.
    [Fact]
    public void ResultEndingInAParagraphMarkPutsTheEndInAParagraphOfItsOwn()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field);

        field.Result.Text = "A\rB\rC\r";

        Assert.Equal("A\rB\rC\r", field.Result.Text);
        Assert.Equal(19, field.Result.Range.Start);
        Assert.Equal(25, field.Result.Range.End);
        Assert.Equal("BeforeA\rB\rC\r\rAfter\r", document.Content.Text);
        Assert.Equal(5, document.Paragraphs.Count);
    }

    //Word: setting "Z" over a result of several paragraphs joins them again: "BeforeZ\rAfter\r".
    [Fact]
    public void ResultWithoutParagraphMarksJoinsTheParagraphsAgain()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field);
        field.Result.Text = "A\rB\rC\r";

        field.Result.Text = "Z";

        Assert.Equal("Z", field.Result.Text);
        Assert.Equal("BeforeZ\rAfter\r", document.Content.Text);
        Assert.Equal(2, document.Paragraphs.Count);
    }

    //Word: "P\rQ\rR" over "A\rB" gives "BeforeP\rQ\rR\rAfter\r".
    [Fact]
    public void ResultOfSeveralParagraphsCanBeReplacedByAnother()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field);
        field.Result.Text = "A\rB";

        field.Result.Text = "P\rQ\rR";

        Assert.Equal("P\rQ\rR", field.Result.Text);
        Assert.Equal("BeforeP\rQ\rR\rAfter\r", document.Content.Text);
    }

    //Word: a field in the middle of "Before" (at 3) gives "BefA\rBore\rAfter\r", the result 16 to 19.
    [Fact]
    public void FieldInTheMiddleOfAParagraphSplitsIt()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field, 3);

        field.Result.Text = "A\rB";

        Assert.Equal("A\rB", field.Result.Text);
        Assert.Equal(16, field.Result.Range.Start);
        Assert.Equal(19, field.Result.Range.End);
        Assert.Equal("BefA\rBore\rAfter\r", document.Content.Text);
    }

    //A field written as one element (w:fldSimple) cannot hold paragraphs, so it becomes a field of begin, code, separator and end first.
    [Fact]
    public void SimpleFieldTakesAResultOfSeveralParagraphs()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Content.Text = "Before\rAfter";
        DocumentFormat.OpenXml.Wordprocessing.Paragraph first = document.Paragraphs[0].Paragraph;
        first.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.SimpleField(
            new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("x")))
        { Instruction = " QUOTE \"x\" " });
        FieldWrapper field = document.Fields[0];

        field.Result.Text = "A\rB\r";

        Assert.Equal("A\rB\r", field.Result.Text);
        Assert.Equal("QUOTE \"x\"", field.Code.Text);
        Assert.Equal("BeforeA\rB\r\rAfter\r", document.Content.Text);
        Assert.Empty(document.WordprocessingDocument.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.SimpleField>());
        Assert.Equal(1, document.Fields.Count);
    }

    //The text keeps the formatting of the first run of the result, in every paragraph, as setting Range.Text does in Word.
    [Fact]
    public void ResultOfSeveralParagraphsKeepsTheFormattingOfTheResult()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field);
        field.Result.Range.Font.Bold = true;

        field.Result.Text = "A\rB";

        Assert.True(field.Result.Range.Font.Bold);
        field.Result.Text = "Z";
        Assert.True(field.Result.Range.Font.Bold);
    }

    //A range after the field stays on its text when the result's paragraphs are joined again.
    [Fact]
    public void RangeAfterTheFieldStaysOnItsTextWhenTheResultIsJoined()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field);
        field.Result.Text = "A\rB\rC\r";
        RangeWrapper after = document.Range(document.Paragraphs[4].Range.Start, document.Paragraphs[4].Range.End);
        Assert.Equal("After\r", after.Text);

        field.Result.Text = "Z";

        Assert.Equal("After\r", after.Text);
        Assert.Equal(22, after.Start);
    }

    //Updating a field whose result is several paragraphs gives it its one-paragraph result again.
    [Fact]
    public void UpdateJoinsAResultOfSeveralParagraphs()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field);
        field.Result.Text = "A\rB\rC\r";

        Assert.True(field.Update());

        Assert.Equal("x", field.Result.Text);
        Assert.Equal("Beforex\rAfter\r", document.Content.Text);
    }

    //Word: the new paragraphs, and the one the end goes to, take the style of the field's paragraph.
    [Fact]
    public void NewParagraphsTakeTheFormattingOfTheFieldsParagraph()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out FieldWrapper field);
        document.Paragraphs[0].Style = "Heading 1";

        field.Result.Text = "A\rB\rC\r";

        Assert.Equal(new[] { "Heading 1", "Heading 1", "Heading 1", "Heading 1", "Normal" }, document.Paragraphs.Select(p => p.Style));
    }
}
