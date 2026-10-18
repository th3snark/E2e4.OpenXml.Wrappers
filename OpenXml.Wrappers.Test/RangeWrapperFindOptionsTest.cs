using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class RangeWrapperFindOptionsTest
{
    private static Body CreateBody(params string[] paragraphs)
    {
        return new Body(paragraphs.Select(t => new Paragraph(new Run(new Text(t) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }))));
    }

    private static string Text(Body body)
    {
        return String.Join("|", body.Elements<Paragraph>().Select(p => String.Concat(p.Descendants<Text>().Select(t => t.Text))));
    }

    [Fact]
    public void SearchingBackwardsFindsTheLastMatchAndCarriesOnTowardsTheStart()
    {
        Body body = CreateBody("one cat", "two cat", "three cat");
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Forward = false;

        Assert.True(range.Find.Execute("cat"));
        Assert.Equal("cat", range.Text);
        Assert.Equal("three cat\r", range.Paragraphs[0].Text);

        Assert.True(range.Find.Execute("cat"));
        Assert.Equal("two cat\r", range.Paragraphs[0].Text);
        Assert.True(range.Find.Execute("cat"));
        Assert.Equal("one cat\r", range.Paragraphs[0].Text);
        Assert.False(range.Find.Execute("cat"));
    }

    [Fact]
    public void SearchingBackwardsReplacesTheLastMatchOnly()
    {
        Body body = CreateBody("cat cat", "cat");

        Assert.True(new RangeWrapper(body).Find.Execute("cat", forward: false, replaceWith: "dog", replace: ReplaceMode.One));

        Assert.Equal("cat cat|dog", Text(body));
    }

    [Fact]
    public void SearchingBackwardsAcrossParagraphMarksFindsTheLastMatch()
    {
        Body body = CreateBody("a", "b", "a", "b");

        Assert.True(new RangeWrapper(body).Find.Execute("a^pb", forward: false, replaceWith: "X", replace: ReplaceMode.One));

        Assert.Equal("a|b|X", Text(body));
    }

    [Fact]
    public void WrapContinuePicksUpMatchesBeforeTheStartOfACollapsedRange()
    {
        Body body = CreateBody("cat here", "nothing");
        Paragraph second = body.Elements<Paragraph>().Last();
        RangeWrapper range = new RangeWrapper(second);
        range.Collapse(CollapseDirection.End);

        Assert.False(new RangeWrapper(second).Find.Execute("cat"));
        Assert.True(new RangeWrapper(second).Find.Execute("cat", wrap: FindWrap.Continue));
    }

    [Fact]
    public void WrapAskIsNotSupported()
    {
        Body body = CreateBody("cat");

        Assert.Throws<NotSupportedException>(() => new RangeWrapper(body).Find.Execute("cat", wrap: FindWrap.Ask));
    }

    [Fact]
    public void SoundsLikeAllWordFormsAndFormatAreNotSupported()
    {
        FindWrapper find = new RangeWrapper(CreateBody("cat")).Find;

        Assert.Throws<NotSupportedException>(() => find.MatchSoundsLike = true);
        Assert.Throws<NotSupportedException>(() => find.MatchAllWordForms = true);
        Assert.Throws<NotSupportedException>(() => find.Format = true);
        Assert.Throws<NotSupportedException>(() => find.Execute("cat", matchSoundsLike: true));
        find.MatchSoundsLike = false;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        Assert.True(find.Execute("cat"));
    }

    [Theory]
    [InlineData("c?t", "cat")]
    [InlineData("c*t", "cat")]
    [InlineData("[bc]at", "cat")]
    [InlineData("[!b]at", "cat")]
    [InlineData("[a-c]at", "cat")]
    [InlineData("<cat>", "cat")]
    [InlineData("ca@t", "cat")]
    [InlineData("ca{1}t", "cat")]
    [InlineData("ca{1,2}t", "cat")]
    [InlineData("ca{,2}t", "cat")]
    [InlineData("c(a)t", "cat")]
    public void WildcardsMatchWhatWordsPatternsDo(string pattern, string expected)
    {
        RangeWrapper range = new RangeWrapper(CreateBody("a dog and a cat sat"));

        Assert.True(range.Find.Execute(pattern, matchWildcards: true));
        Assert.Equal(expected, range.Text);
    }

    [Fact]
    public void WildcardsDoNotMatchWhatTheyShouldNot()
    {
        RangeWrapper range = new RangeWrapper(CreateBody("concatenate"));

        Assert.False(range.Find.Execute("<cat>", matchWildcards: true));
        Assert.False(new RangeWrapper(CreateBody("cat")).Find.Execute("[!c]at", matchWildcards: true));
        Assert.False(new RangeWrapper(CreateBody("cat")).Find.Execute("CAT", matchCase: true, matchWildcards: true));
        Assert.True(new RangeWrapper(CreateBody("cat")).Find.Execute("CAT", matchWildcards: true));
    }

    [Fact]
    public void ABackslashFindsACharacterThatHasAMeaning()
    {
        Assert.True(new RangeWrapper(CreateBody("what? really")).Find.Execute("what\\?", matchWildcards: true));
        Assert.False(new RangeWrapper(CreateBody("whatx really")).Find.Execute("what\\?", matchWildcards: true));
    }

    [Fact]
    public void GroupsCanBeUsedInTheReplacement()
    {
        Body body = CreateBody("Smith, John and Jones, Mary");

        new RangeWrapper(body).Find.Execute("(<[A-Za-z]@>), (<[A-Za-z]@>)", matchWildcards: true, replaceWith: "\\2 \\1", replace: ReplaceMode.All);

        Assert.Equal("John Smith and Mary Jones", Text(body));
    }

    [Fact]
    public void AGroupCanBeRepeatedInThePattern()
    {
        Body body = CreateBody("this is is a test and a a mistake");

        new RangeWrapper(body).Find.Execute("(<[a-z]@>) \\1", matchWildcards: true, replaceWith: "\\1", replace: ReplaceMode.All);

        Assert.Equal("this is a test and a mistake", Text(body));
    }

    [Fact]
    public void AWildcardMatchCanRunAcrossParagraphs()
    {
        Body body = CreateBody("start", "middle", "end");

        Assert.True(new RangeWrapper(body).Find.Execute("start*end", matchWildcards: true, replaceWith: "X", replace: ReplaceMode.All));

        Assert.Equal("X", Text(body));
    }

    [Fact]
    public void AnInvalidPatternThrows()
    {
        RangeWrapper range = new RangeWrapper(CreateBody("cat"));

        Assert.Throws<ArgumentException>(() => range.Find.Execute("[abc", matchWildcards: true));
        Assert.Throws<ArgumentException>(() => range.Find.Execute("ca{2", matchWildcards: true));
        Assert.Throws<ArgumentException>(() => range.Find.Execute("@cat", matchWildcards: true));
    }

    [Fact]
    public void TheFoundTextCanBeUsedInTheReplacement()
    {
        Body body = CreateBody("one cat and two cats");

        new RangeWrapper(body).Find.Execute("cat", replaceWith: "[^&]", replace: ReplaceMode.All);

        Assert.Equal("one [cat] and two [cat]s", Text(body));
    }

    [Fact]
    public void CodesForKindsOfCharacterFindThem()
    {
        Assert.True(new RangeWrapper(CreateBody("a  b")).Find.Execute("a^wb"));
        Assert.True(new RangeWrapper(CreateBody("room 42b")).Find.Execute("^#^#b"));
        Assert.True(new RangeWrapper(CreateBody("x7y")).Find.Execute("x^#y"));
        Assert.False(new RangeWrapper(CreateBody("xzy")).Find.Execute("x^#y"));
        Assert.True(new RangeWrapper(CreateBody("a1")).Find.Execute("^$^#"));
        Assert.True(new RangeWrapper(CreateBody("a1")).Find.Execute("a^?"));
    }

    [Fact]
    public void MoreCodesStandForCharacters()
    {
        Body body = CreateBody("a b — c");

        Assert.True(new RangeWrapper(body).Find.Execute("a^sb"));
        new RangeWrapper(body).Find.Execute("^+", replaceWith: "^=", replace: ReplaceMode.All);

        Assert.Equal("a b – c", Text(body));
    }

    [Fact]
    public void ExecuteTakesWordsParametersInWordsOrder()
    {
        Body body = CreateBody("Cat cat");

        Assert.True(new RangeWrapper(body).Find.Execute("Cat", true, false, false, false, false, true, FindWrap.Stop, false, "Dog", ReplaceMode.All));

        Assert.Equal("Dog cat", Text(body));
    }
}
