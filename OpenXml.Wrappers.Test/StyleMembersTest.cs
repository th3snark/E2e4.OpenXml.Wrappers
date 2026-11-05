using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The members of <see cref="StyleWrapper"/> that are flags and links, the language, and the base and next styles. The markup is what Word 16
/// wrote (the style-* scenarios of MarkupConformanceTest).
/// </summary>
public class StyleMembersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public StyleMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ANewParagraphStyle_HasWordsDefaults_AndAStyleLinkedToIt()
    {
        StyleWrapper style = _document.Styles.Add("My Style", StyleType.Paragraph);

        Assert.False(style.QuickStyle);
        Assert.False(style.Locked);
        Assert.False(style.UnhideWhenUsed);
        Assert.False(style.AutomaticallyUpdate);
        Assert.False(style.Visibility);
        Assert.False(style.NoSpaceBetweenParagraphsOfSameStyle);
        Assert.True(style.InUse);
        Assert.Equal(1, style.Priority);
        Assert.True(style.Linked);
        Assert.Equal("My Style Char", style.LinkStyle.Name);
        Assert.Equal(StyleType.Character, style.LinkStyle.Type);
        Assert.Equal("My Style", style.LinkStyle.LinkStyle.Name);
        Assert.Equal("Default Paragraph Font", style.LinkStyle.BaseStyle);
    }

    [Fact]
    public void ACharacterStyle_IsNotLinked_AndIsBasedOnTheDefault()
    {
        StyleWrapper style = _document.Styles.Add("Mine", StyleType.Character);

        Assert.False(style.Linked);
        Assert.Null(style.LinkStyle);
        Assert.Equal("Default Paragraph Font", style.BaseStyle);
    }

    [Fact]
    public void Flags_ReadWhatTheyAreSetTo_AndTurnOff()
    {
        StyleWrapper style = _document.Styles.Add("Flags", StyleType.Paragraph);

        style.QuickStyle = true;
        style.Locked = true;
        style.UnhideWhenUsed = true;
        style.AutomaticallyUpdate = true;
        style.Visibility = true;
        style.NoSpaceBetweenParagraphsOfSameStyle = true;

        Assert.True(style.QuickStyle && style.Locked && style.UnhideWhenUsed && style.AutomaticallyUpdate && style.Visibility && style.NoSpaceBetweenParagraphsOfSameStyle);
        Assert.NotNull(style.Style.StyleParagraphProperties.ContextualSpacing);

        style.QuickStyle = false;
        style.NoSpaceBetweenParagraphsOfSameStyle = false;
        Assert.False(style.QuickStyle);
        Assert.False(style.NoSpaceBetweenParagraphsOfSameStyle);
        Assert.Null(style.Style.StyleParagraphProperties);
    }

    [Fact]
    public void Priority_IsOneMoreThanWhatIsStored_AndTheLinkedStyleFollows()
    {
        StyleWrapper style = _document.Styles.Add("Prioritised", StyleType.Paragraph);

        style.Priority = 5;

        Assert.Equal(4, style.Style.UIPriority.Val.Value);
        Assert.Equal(5, style.LinkStyle.Priority);

        style.Priority = 1;
        Assert.Null(style.Style.UIPriority);
        Assert.Throws<ArgumentOutOfRangeException>(() => style.Priority = 0);
    }

    [Fact]
    public void BuiltInStyles_ReadTheirPriority()
    {
        StyleWrapper heading = _document.Styles["Heading 1"];

        Assert.Equal(10, heading.Priority);
        Assert.True(heading.QuickStyle);
    }

    [Fact]
    public void BaseStyle_AndNextParagraphStyle_CanBeSetAndRemoved()
    {
        StyleWrapper style = _document.Styles.Add("Based", StyleType.Paragraph);

        style.BaseStyle = "Heading 1";
        style.NextParagraphStyle = "Heading 2";

        Assert.Equal("Heading 1", style.BaseStyle);
        Assert.Equal("Heading 2", style.NextParagraphStyle);
        Assert.NotNull(style.LinkStyle.Style.StyleRunProperties);

        style.BaseStyle = null;
        style.NextParagraphStyle = null;
        Assert.Null(style.BaseStyle);
        Assert.Null(style.NextParagraphStyle);
        Assert.Throws<ArgumentException>(() => style.BaseStyle = "No Such Style");
    }

    [Fact]
    public void Language_IsSetOnTheStyleAndItsLinkedStyle()
    {
        StyleWrapper style = _document.Styles.Add("Lang", StyleType.Paragraph);

        style.LanguageID = 1033;
        style.LanguageIDFarEast = 1041;
        style.NoProofing = true;

        Assert.Equal(1033, style.LanguageID);
        Assert.Equal(1041, style.LanguageIDFarEast);
        Assert.True(style.NoProofing);
        Assert.Equal(1033, style.LinkStyle.LanguageID);
        Assert.True(style.LinkStyle.NoProofing);

        style.LanguageID = 0;
        style.LanguageIDFarEast = 0;
        style.NoProofing = false;
        Assert.Equal(0, style.LanguageID);
        Assert.True(style.Style.StyleRunProperties == null, style.Style.OuterXml);
    }

    [Fact]
    public void Language_Rejects_ANumberThatIsNotALanguage()
    {
        StyleWrapper style = _document.Styles.Add("Lang", StyleType.Paragraph);

        Assert.Throws<ArgumentOutOfRangeException>(() => style.LanguageID = 39321);
    }

    [Fact]
    public void DeletingAParagraphStyle_UnlinksTheCharacterStyle()
    {
        StyleWrapper style = _document.Styles.Add("Gone", StyleType.Paragraph);
        StyleWrapper linked = style.LinkStyle;

        style.Delete();

        Assert.False(linked.Linked);
    }
}
