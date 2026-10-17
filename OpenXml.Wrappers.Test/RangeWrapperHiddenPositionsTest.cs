using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Positions as Word counts them: the begin, separator, end and code of a field, and pictures and references, take up positions but show no text.
/// </summary>
public class RangeWrapperHiddenPositionsTest
{
    private static Run Mark(FieldCharValues type)
    {
        return new Run(new FieldChar() { FieldCharType = type });
    }

    private static Run Code(string code)
    {
        return new Run(new FieldCode(code) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Run Plain(string text)
    {
        return new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    //"ab", a field with the code " PAGE " and the result "12", and "cd".
    private static Body BodyWithField()
    {
        return new Body(new Paragraph(Plain("ab"), Mark(FieldCharValues.Begin), Code(" PAGE "), Mark(FieldCharValues.Separate), Plain("12"), Mark(FieldCharValues.End), Plain("cd")));
    }

    private static RangeWrapper Range(Body body, int start, int end)
    {
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(start, end);
        return range;
    }

    [Fact]
    public void AFieldTakesUpPositionsForItsMarksAndItsCode()
    {
        Body body = BodyWithField();

        //2 + begin 1 + code 6 + separate 1 + result 2 + end 1 + 2 + paragraph mark 1.
        Assert.Equal(16, new RangeWrapper(body).StoryLength);
        Assert.Equal("ab12cd\r", new RangeWrapper(body).Text);
        Assert.Equal("ab", Range(body, 0, 2).Text);
        Assert.Equal("12", Range(body, 10, 12).Text);
        Assert.Equal("cd", Range(body, 13, 15).Text);
        Assert.Equal(string.Empty, Range(body, 3, 9).Text);
    }

    [Fact]
    public void ARangeOfTheFieldElementsHasThePositionsOfTheWholeField()
    {
        Body body = BodyWithField();
        Paragraph paragraph = body.Elements<Paragraph>().Single();

        RangeWrapper range = new ParagraphWrapper(paragraph).Range;

        Assert.Equal(0, range.Start);
        Assert.Equal(16, range.End);
        Assert.Equal(16, range.Characters.Count);
        Assert.Equal("ab12cd\r", string.Concat(range.Characters.Select(c => c.Text)));
    }

    [Fact]
    public void WordsDoNotBreakAtTheHiddenPositionsOfAField()
    {
        Body body = BodyWithField();

        RangeWrapper all = new RangeWrapper(body);

        Assert.Equal(new[] { "ab12cd", "\r" }, all.Words.Select(w => w.Text).ToArray());
        Assert.Equal(16, all.Words.Sum(w => w.End - w.Start));
    }

    [Fact]
    public void FormattingInsideTheCodeOfAFieldLeavesAWholeField()
    {
        Body body = BodyWithField();

        Range(body, 4, 7).Font.Bold = true;

        Assert.Equal("ab12cd\r", new RangeWrapper(body).Text);
        Assert.Equal(16, new RangeWrapper(body).StoryLength);
        Assert.Equal(3, body.Descendants<FieldChar>().Count());
        Assert.Equal(" PAGE ", string.Concat(body.Descendants<FieldCode>().Select(c => c.Text)));
    }

    [Fact]
    public void ARangeStartingInTheCodeAndDeletedTakesTheWholeField()
    {
        Body body = BodyWithField();

        Range(body, 5, 11).Delete();

        Assert.Equal("abcd\r", new RangeWrapper(body).Text);
        Assert.Empty(body.Descendants<FieldChar>());
    }

    [Fact]
    public void AFindCanMatchAcrossAFieldResult()
    {
        Body body = new Body(new Paragraph(Plain("Page "), Mark(FieldCharValues.Begin), Code(" PAGE "), Mark(FieldCharValues.Separate), Plain("3"), Mark(FieldCharValues.End), Plain(" of 9")));

        RangeWrapper range = new RangeWrapper(body);
        Assert.True(range.Find.Execute("Page 3 of"));
        Assert.Equal("Page 3 of", range.Text);

        range = new RangeWrapper(body);
        Assert.True(range.Find.Execute("Page 3", replaceWith: "X", replace: ReplaceMode.All));
        Assert.Equal("X of 9\r", new RangeWrapper(body).Text);
        Assert.Empty(body.Descendants<FieldChar>());
    }

    [Fact]
    public void APictureAndReferencesTakeUpOnePositionEach()
    {
        Paragraph paragraph = new Paragraph(Plain("a"), new Run(new Drawing()), Plain("b"), new Run(new FootnoteReference() { Id = 2 }), Plain("c"));
        Body body = new Body(paragraph);

        //a, picture, b, reference, c, paragraph mark.
        Assert.Equal(6, new RangeWrapper(body).StoryLength);
        Assert.Equal("b", Range(body, 2, 3).Text);
        Assert.Equal(string.Empty, Range(body, 1, 2).Text);
        Assert.Equal("abc\r", new RangeWrapper(body).Text);
        Assert.Equal(4, Range(body, 4, 5).Start);
    }

    [Fact]
    public void FormattingAndDeletingAroundAPictureWorks()
    {
        Paragraph paragraph = new Paragraph(Plain("one"), new Run(new Drawing()), Plain("two"));
        Body body = new Body(paragraph);

        Range(body, 2, 5).Font.Italic = true;
        Assert.True(Range(body, 2, 5).Font.Italic);
        Assert.Equal("onetwo\r", new RangeWrapper(body).Text);

        Range(body, 3, 4).Delete();
        Assert.Empty(body.Descendants<Drawing>());
        Assert.Equal("onetwo\r", new RangeWrapper(body).Text);
        Assert.Equal(7, new RangeWrapper(body).StoryLength);
    }

    [Fact]
    public void ARangeFollowsTextAcrossAField()
    {
        Body body = BodyWithField();
        RangeWrapper cd = Range(body, 13, 15);

        Range(body, 0, 0).InsertBefore(">>");

        Assert.Equal("cd", cd.Text);
        Assert.Equal(15, cd.Start);
    }
}
