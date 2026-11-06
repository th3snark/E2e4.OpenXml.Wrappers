using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The word edge, punctuation and white space options of <c>Find</c>, and the language and proofing of the formatting it looks for. The
/// results are what Word gave (the find-match-* and find-ignore-* scenarios of MarkupConformanceTest).
/// </summary>
public class RangeWrapperFindMatchOptionsTest
{
    private static Body CreateBody(string text)
    {
        return new Body(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
    }

    private static string Replaced(string text, Action<FindWrapper> options, string find)
    {
        Body body = CreateBody(text);
        FindWrapper wrapper = new RangeWrapper(body).Find;
        options(wrapper);
        wrapper.Execute(find, replaceWith: "X", replace: ReplaceMode.All);
        return string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }

    [Fact]
    public void APrefixStartsAWord()
    {
        Assert.Equal("X Xfix Xach unpre repre", Replaced("pre prefix preach unpre repre", f => f.MatchPrefix = true, "pre"));
    }

    [Fact]
    public void ASuffixEndsAWord()
    {
        Assert.Equal("X sX singX ingot", Replaced("ing sing singing ingot", f => f.MatchSuffix = true, "ing"));
    }

    [Fact]
    public void APrefixAndASuffixIsAWholeWord()
    {
        Assert.Equal("X concat cats X", Replaced("cat concat cats cat", f => { f.MatchPrefix = true; f.MatchSuffix = true; }, "cat"));
        Assert.Equal("X concat cats X", Replaced("cat concat cats cat", f => f.MatchWholeWord = true, "cat"));
    }

    [Fact]
    public void PunctuationCanBeIgnored()
    {
        Assert.Equal("X X X co op", Replaced("co-op coop co,op co op", f => f.IgnorePunct = true, "coop"));
        Assert.Equal("X X co op", Replaced("co-op coop co op", f => f.IgnorePunct = true, "co-op"));
    }

    [Fact]
    public void WhiteSpaceCanBeIgnored()
    {
        Assert.Equal("X X X X", Replaced("a b ab a  b ab", f => f.IgnoreSpace = true, "ab"));
    }

    [Fact]
    public void AMatchThatIgnoresPunctuationEndsAtItsLastLetter()
    {
        Body body = CreateBody("co-op, then");
        RangeWrapper range = new RangeWrapper(body);
        range.Find.IgnorePunct = true;

        Assert.True(range.Find.Execute("coop"));
        Assert.Equal("co-op", range.Text);
    }

    [Fact]
    public void TextThatIsOnlyPunctuationIsNotFoundWhenPunctuationIsIgnored()
    {
        FindWrapper find = new RangeWrapper(CreateBody("a-b")).Find;
        find.IgnorePunct = true;

        Assert.False(find.Execute("-"));
    }

    [Fact]
    public void ALanguageAndNoProofingAreGivenToTheReplacement()
    {
        Body body = CreateBody("One two three");
        FindWrapper find = new RangeWrapper(body).Find;
        find.Replacement.LanguageID = 1036;
        find.Replacement.NoProofing = true;

        Assert.True(find.Execute("two", replaceWith: "^&", replace: ReplaceMode.All));

        Run run = body.Descendants<Run>().Single(r => r.InnerText == "two");
        Assert.Equal("fr-FR", run.RunProperties.Languages.Val.Value);
        Assert.NotNull(run.RunProperties.NoProof);
        Assert.Equal(1036, find.Replacement.LanguageID);
        Assert.True(find.Replacement.NoProofing);
    }

    [Fact]
    public void ALanguageIsFoundAndNoProofingToo()
    {
        Body body = new Body(new Paragraph(
            new Run(new Text("one ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new RunProperties(new NoProof(), new Languages() { Val = "fr-FR" }), new Text("two"))));
        RangeWrapper range = new RangeWrapper(body); range.Find.LanguageID = 1036; Assert.True(range.Find.Execute(string.Empty)); Assert.Equal("two", range.Text); RangeWrapper other = new RangeWrapper(body); other.Find.NoProofing = false; Assert.True(other.Find.Execute(string.Empty)); Assert.Equal("one ", other.Text);
    }

    [Fact]
    public void ALanguageThatIsNotOneIsRejected()
    {
        FindWrapper find = new RangeWrapper(CreateBody("x")).Find;

        Assert.Throws<ArgumentOutOfRangeException>(() => find.LanguageID = 39321);
        Assert.Throws<ArgumentOutOfRangeException>(() => find.Replacement.LanguageIDFarEast = 39321);
    }
}
