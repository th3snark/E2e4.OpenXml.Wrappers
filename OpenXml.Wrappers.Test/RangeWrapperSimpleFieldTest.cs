using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Fields written as one element (<c>w:fldSimple</c>) take up positions as Word counts them (checked against Word): begin, code, separator,
/// result and end.
/// </summary>
public class RangeWrapperSimpleFieldTest
{
    private static Run Plain(string text)
    {
        return new Run(new Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
    }

    private static Body BodyWithField()
    {
        return new Body(new Paragraph(Plain("ab"), new SimpleField(Plain("7")) { Instruction = " PAGE " }, Plain("cd")));
    }

    [Fact]
    public void PositionsAreTheSameAsInWord()
    {
        Body body = BodyWithField();

        //ab at 0-2, begin at 2, " PAGE " at 3-9, separator at 9, 7 at 10, end at 11, cd at 12-14, the end of the paragraph at 14.
        RangeWrapper content = new RangeWrapper(body);
        Assert.Equal(15, content.End);

        RangeWrapper cd = new RangeWrapper(body);
        Assert.True(cd.Find.Execute("cd"));
        Assert.Equal(12, cd.Start);
        Assert.Equal(14, cd.End);

        RangeWrapper result = new RangeWrapper(body);
        result.SetRange(10, 11);
        Assert.Equal("7", result.Text);
    }

    [Fact]
    public void TheTextLeavesOutTheCodeAndMarks()
    {
        Assert.Equal("ab7cd\r", new RangeWrapper(BodyWithField()).Text);
    }

    [Fact]
    public void DeletingPartOfAFieldDeletesAllOfIt()
    {
        Body body = BodyWithField();
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(5, 6);

        range.Delete();

        Assert.Empty(body.Descendants<SimpleField>());
        Assert.Equal("abcd\r", new RangeWrapper(body).Text);
    }

    [Fact]
    public void ARangeThatKeepsItsPlaceFollowsTextAfterTheField()
    {
        Body body = BodyWithField();
        RangeWrapper cd = new RangeWrapper(body);
        cd.SetRange(12, 14);

        RangeWrapper before = new RangeWrapper(body);
        before.SetRange(0, 2);
        before.Delete();

        Assert.Equal("cd", cd.Text);
        Assert.Equal(10, cd.Start);
    }

    [Fact]
    public void InsertingBeforeAndAfterAFieldDoesNotChangeIt()
    {
        Body body = BodyWithField();
        RangeWrapper after = new RangeWrapper(body);
        after.SetRange(12, 12);

        after.InsertAfter("X");

        Assert.Single(body.Descendants<SimpleField>());
        Assert.Equal("ab7Xcd\r", new RangeWrapper(body).Text);
        Assert.Equal("7", body.Descendants<SimpleField>().Single().InnerText);
    }
}
