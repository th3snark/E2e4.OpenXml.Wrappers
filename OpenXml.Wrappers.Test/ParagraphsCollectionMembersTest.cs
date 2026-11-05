using System;
using System.IO;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The members of <c>Paragraphs</c> that apply the paragraph format to every paragraph, spacing steps, outline moves and
/// <c>Paragraphs.Add(Range)</c>. The results are those Word 16 gave for the same content.
/// </summary>
public class ParagraphsCollectionMembersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ParagraphsCollectionMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "A\rB\rC\rD\rE";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private ParagraphWrapper Paragraph(int index)
    {
        return _document.Content.Paragraphs[index];
    }

    private string Outline()
    {
        return string.Join(" ", _document.Content.Paragraphs.Select(p => p.Style + "/" + (int)p.Format.OutlineLevel.Value));
    }

    private void SetStyles(params string[] styles)
    {
        for (int i = 0; i < styles.Length; i++)
        {
            if (styles[i] != null)
            {
                Paragraph(i).Style = styles[i];
            }
        }
    }

    [Fact]
    public void IncreaseAndDecreaseSpacing_StepSixPoints_NeverBelowZero()
    {
        ParagraphWrapperCollection paragraphs = _document.Content.Paragraphs;
        double after = paragraphs.SpaceAfter.Value;

        paragraphs.IncreaseSpacing();
        Assert.Equal(6, paragraphs.SpaceBefore);
        Assert.Equal(after + 6, paragraphs.SpaceAfter);

        paragraphs.IncreaseSpacing();
        Assert.Equal(12, paragraphs.SpaceBefore);

        paragraphs.DecreaseSpacing();
        paragraphs.DecreaseSpacing();
        paragraphs.DecreaseSpacing();
        Assert.Equal(0, paragraphs.SpaceBefore);
        Assert.Equal(Math.Max(0, after - 6), paragraphs.SpaceAfter);

        paragraphs.DecreaseSpacing();
        paragraphs.DecreaseSpacing();
        Assert.Equal(0, paragraphs.SpaceAfter);
    }

    [Fact]
    public void SpacingSteps_KeepWhatEachParagraphHas()
    {
        Paragraph(1).Format.SpaceBefore = 10;

        _document.Content.Paragraphs.IncreaseSpacing();

        Assert.Equal(6, Paragraph(0).Format.SpaceBefore);
        Assert.Equal(16, Paragraph(1).Format.SpaceBefore);
    }

    [Fact]
    public void FormatMembers_ReadNullWhereTheParagraphsDiffer_AndSetEveryOne()
    {
        ParagraphWrapperCollection paragraphs = _document.Content.Paragraphs;
        Paragraph(0).Format.KeepWithNext = true;

        Assert.Null(paragraphs.KeepWithNext);

        paragraphs.KeepWithNext = true;
        paragraphs.Alignment = ParagraphAlignment.Center;
        paragraphs.LeftIndent = 18;
        paragraphs.WordWrap = false;

        Assert.True(paragraphs.KeepWithNext);
        Assert.All(paragraphs, p => Assert.Equal(ParagraphAlignment.Center, p.Alignment));
        Assert.All(paragraphs, p => Assert.Equal(18, p.LeftIndent));
        Assert.False(paragraphs.WordWrap);
    }

    [Fact]
    public void Shortcuts_ApplyToEveryParagraph()
    {
        ParagraphWrapperCollection paragraphs = _document.Content.Paragraphs;

        paragraphs.OpenUp();
        Assert.Equal(12, paragraphs.SpaceBefore);
        paragraphs.Space15();
        Assert.Equal(LineSpacingRule.OnePointFive, paragraphs.LineSpacingRule);
        paragraphs.Indent();
        Assert.Equal(36, paragraphs.LeftIndent);
        paragraphs.Indent();
        paragraphs.Outdent();
        Assert.Equal(36, paragraphs.LeftIndent);
        paragraphs.Reset();
        Assert.Equal(0, paragraphs.SpaceBefore);
        Assert.Equal(0, paragraphs.LeftIndent);
    }

    [Fact]
    public void BordersAndShading_ApplyToEveryParagraph()
    {
        ParagraphWrapperCollection paragraphs = _document.Content.Paragraphs;

        paragraphs.Borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.SingleLine;
        paragraphs.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);

        Assert.All(paragraphs, p => Assert.Equal(BorderLineStyle.SingleLine, p.Borders[BorderPosition.Bottom].LineStyle));
        Assert.All(paragraphs, p => Assert.Equal(ColorHelper.RGB(255, 0, 0), (int)p.Shading.BackgroundPatternColor.Value.Value));
    }

    [Theory]
    [InlineData("Heading 1", "Heading 2")]
    [InlineData("Heading 2", "Heading 3")]
    [InlineData("Heading 8", "Heading 9")]
    [InlineData("Heading 9", "Heading 9")]
    public void OutlineDemote_MovesAHeadingDown(string style, string expected)
    {
        Paragraph(0).Style = style;

        Paragraph(0).OutlineDemote();

        Assert.Equal(expected, Paragraph(0).Style);
    }

    [Theory]
    [InlineData("Heading 1", "Heading 1")]
    [InlineData("Heading 2", "Heading 1")]
    [InlineData("Heading 9", "Heading 8")]
    public void OutlinePromote_MovesAHeadingUp(string style, string expected)
    {
        Paragraph(0).Style = style;

        Paragraph(0).OutlinePromote();

        Assert.Equal(expected, Paragraph(0).Style);
    }

    [Fact]
    public void BodyText_TakesTheLevelOfTheHeadingBefore()
    {
        SetStyles("Heading 3", null, null);

        Paragraph(2).OutlineDemote();
        Assert.Equal("Heading 4", Paragraph(2).Style);

        Paragraph(1).OutlinePromote();
        Assert.Equal("Heading 3", Paragraph(1).Style);
    }

    [Fact]
    public void BodyTextWithNoHeadingBefore_IsDemotedToHeading2_AndPromotedToHeading1()
    {
        Paragraph(1).OutlinePromote();
        Assert.Equal("Heading 1", Paragraph(1).Style);

        Paragraph(0).OutlineDemote();
        Assert.Equal("Heading 2", Paragraph(0).Style);
    }

    [Fact]
    public void BodyText_FollowsTheNearestHeadingBefore_NotTheNeighbour()
    {
        SetStyles("Heading 2", "Heading 5", null);

        Paragraph(2).OutlineDemote();

        Assert.Equal("Heading 6", Paragraph(2).Style);
    }

    [Fact]
    public void ALevelOfItsOwn_IsMovedWithoutChangingTheStyle()
    {
        Paragraph(0).Format.OutlineLevel = ParagraphOutlineLevel.Level3;

        Paragraph(0).OutlineDemote();
        Assert.Equal(ParagraphOutlineLevel.Level4, Paragraph(0).Format.OutlineLevel);
        Assert.Equal("Normal", Paragraph(0).Style);

        Paragraph(0).OutlinePromote();
        Paragraph(0).OutlinePromote();
        Paragraph(0).OutlinePromote();
        Paragraph(0).OutlinePromote();
        Assert.Equal(ParagraphOutlineLevel.Level1, Paragraph(0).Format.OutlineLevel);
    }

    [Fact]
    public void OutlineDemoteToBody_MakesNormalAndDropsALevelOfItsOwn()
    {
        SetStyles("Heading 2");
        Paragraph(1).Format.OutlineLevel = ParagraphOutlineLevel.Level3;

        Paragraph(0).OutlineDemoteToBody();
        Paragraph(1).OutlineDemoteToBody();

        Assert.Equal("Normal/10 Normal/10 Normal/10 Normal/10 Normal/10", Outline());
    }

    [Fact]
    public void Outline_OnTheCollection_MovesEveryParagraph()
    {
        SetStyles("Heading 1", "Heading 2", "Heading 3");

        _document.Range(0, 6).Paragraphs.OutlineDemote();
        Assert.StartsWith("Heading 2/2 Heading 3/3 Heading 4/4", Outline(), StringComparison.Ordinal);

        _document.Range(0, 6).Paragraphs.OutlineDemoteToBody();
        Assert.StartsWith("Normal/10 Normal/10 Normal/10", Outline(), StringComparison.Ordinal);
    }

    [Fact]
    public void Direct_FormattingOfAHeading_IsKeptWhenItMoves()
    {
        SetStyles("Heading 2");
        Paragraph(0).Range.Font.Bold = true;

        Paragraph(0).OutlineDemote();

        Assert.Equal("Heading 3", Paragraph(0).Style);
        Assert.True(Paragraph(0).Range.Font.Bold);
    }

    [Fact]
    public void Add_OfARange_PutsAnEmptyParagraphBeforeTheParagraphItIsIn()
    {
        ParagraphWrapper added = _document.Content.Paragraphs.Add(Paragraph(1).Range);

        Assert.Equal(6, _document.Content.Paragraphs.Count);
        Assert.Equal("\r", added.Text);
        Assert.Equal("A\r\rB\rC\rD\rE\r", _document.Content.Text);
        Assert.Equal(added, Paragraph(1));
    }

    [Fact]
    public void Add_OfACollapsedRangeInText_SplitsTheParagraph()
    {
        _document.Content.Text = "Body\rNext";

        ParagraphWrapper added = _document.Content.Paragraphs.Add(_document.Range(3, 3));

        Assert.Equal("Bod\r", added.Text);
        Assert.Equal("Bod\ry\rNext\r", _document.Content.Text);
        Assert.Equal(3, _document.Content.Paragraphs.Count);
    }

    [Fact]
    public void Add_WithNoRange_AddsAtTheEnd()
    {
        ParagraphWrapper added = _document.Content.Paragraphs.Add();

        Assert.Equal(6, _document.Content.Paragraphs.Count);
        Assert.Equal(added, _document.Content.Paragraphs.Last);
    }

    [Fact]
    public void Add_RejectsNull()
    {
        Assert.Throws<ArgumentNullException>(() => _document.Content.Paragraphs.Add(null));
    }
}
