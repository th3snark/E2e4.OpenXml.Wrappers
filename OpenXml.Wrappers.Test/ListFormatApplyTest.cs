using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Making, changing and removing lists. The markup of each operation is Word 16's (the list-* scenarios of MarkupConformanceTest); these tests
/// cover what a scenario with one list cannot: several lists, and what is checked by value.
/// </summary>
public class ListFormatApplyTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ListFormatApplyTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("One\rTwo\rThree\rFour");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    private ListFormatWrapper Format(int paragraph)
    {
        return _document.Paragraphs[paragraph].Range.ListFormat;
    }

    [Fact]
    public void ApplyingTheBulletTwiceReusesTheList()
    {
        Format(0).ApplyBulletDefault();
        Format(2).ApplyBulletDefault();

        Assert.Equal(1, _document.Lists.Count);
        Assert.Equal(1, _document.ListTemplates.Count);
        Assert.Equal(2, _document.Lists[0].CountNumberedItems());
        Assert.Equal(ListType.Bullet, Format(2).ListType);
        AssertValid();
    }

    [Fact]
    public void NumbersGoOnFromTheParagraphBeforeAndStartAgainAfterAGap()
    {
        Format(0).ApplyNumberDefault();
        Format(1).ApplyNumberDefault();
        Format(3).ApplyNumberDefault();

        Assert.Equal("2.", Format(1).ListString);
        Assert.Equal("1.", Format(3).ListString);
        Assert.Equal(2, _document.Lists.Count);
        Assert.Equal(2, _document.ListTemplates.Count);
        AssertValid();
    }

    [Fact]
    public void AListPutsTheListParagraphStyleOnPlainParagraphsOnly()
    {
        _document.Paragraphs[0].Range.Style = "Heading 1";
        Format(0).ApplyNumberDefault();
        Format(1).ApplyNumberDefault();

        Assert.Equal("Heading1", _document.Paragraphs[0].Paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
        Assert.Equal("ListParagraph", _document.Paragraphs[1].Paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
        AssertValid();
    }

    [Fact]
    public void IndentingStopsAtTheLastLevelAndOutdentingAtTheFirst()
    {
        _document.Content.ListFormat.ApplyOutlineNumberDefault();

        for (int i = 0; i < 12; i++)
        {
            Format(1).ListIndent();
        }
        Format(0).ListOutdent();

        Assert.Equal(9, Format(1).ListLevelNumber);
        Assert.Equal(1, Format(0).ListLevelNumber);
        AssertValid();
    }

    [Fact]
    public void ASingleLevelListDoesNotIndent()
    {
        _document.Content.ListFormat.ApplyNumberDefault();

        Format(1).ListIndent();

        Assert.Equal(1, Format(1).ListLevelNumber);
    }

    [Fact]
    public void TheLevelNumberCanBeSetButOnlyToALevelTheListHas()
    {
        _document.Content.ListFormat.ApplyOutlineNumberDefault();

        Format(2).ListLevelNumber = 4;

        Assert.Equal(4, Format(2).ListLevelNumber);
        Assert.Equal("(1)", Format(2).ListString);
        Assert.Throws<ArgumentOutOfRangeException>(() => Format(2).ListLevelNumber = 10);
        Assert.Throws<ArgumentOutOfRangeException>(() => Format(2).ListLevelNumber = 0);
    }

    [Fact]
    public void RemovingNumbersLeavesAParagraphThatIsNotInAList()
    {
        _document.Content.ListFormat.ApplyBulletDefault();

        Format(1).RemoveNumbers();

        Assert.Equal(ListType.NoNumbering, Format(1).ListType);
        Assert.Equal(3, _document.ListParagraphs.Count);
        Assert.Null(_document.Paragraphs[1].Paragraph.ParagraphProperties);
        AssertValid();
    }

    [Fact]
    public void ConvertingNumbersLeavesTheirTextAndTheIndent()
    {
        _document.Content.ListFormat.ApplyNumberDefault();

        Format(0).ConvertNumbersToText();

        Assert.Equal("1.\tOne", _document.Paragraphs[0].Range.Text.TrimEnd('\r'));
        Assert.Equal(ListType.NoNumbering, Format(0).ListType);
        Assert.Equal("1.", Format(1).ListString);
        Indentation indentation = _document.Paragraphs[0].Paragraph.ParagraphProperties.Indentation;
        Assert.Equal("720", indentation.Left?.Value ?? indentation.Start?.Value);
        AssertValid();
    }

    [Fact]
    public void ATemplateOfTheDocumentCanBeAppliedAgain()
    {
        Format(0).ApplyNumberDefault();
        ListTemplateWrapper template = _document.ListTemplates[0];

        Format(2).ApplyListTemplate(template);
        Format(1).ApplyListTemplate(template, true);

        Assert.Equal("2.", Format(1).ListString);
        Assert.Equal("1.", Format(2).ListString);
        Assert.Equal(2, _document.Lists.Count);
        AssertValid();
    }

    [Fact]
    public void ATemplateOfAnotherDocumentIsRejected()
    {
        string otherPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
        try
        {
            using (WordprocessingDocumentWrapper other = WordprocessingDocumentWrapper.Create(otherPath, true, false))
            {
                other.Content.InsertAfter("x");
                other.Paragraphs[0].Range.ListFormat.ApplyBulletDefault();
                ListTemplateWrapper foreign = other.ListTemplates[0];

                Assert.Throws<ArgumentException>(() => Format(0).ApplyListTemplate(foreign));
            }
        }
        finally
        {
            System.IO.File.Delete(otherPath);
        }
    }

    [Fact]
    public void ANewListDoesNotNeedANumberingPartInTheDocumentBefore()
    {
        Assert.Null(_document.WordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart);

        Format(0).ApplyBulletDefault();

        Assert.NotNull(_document.WordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart);
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Elements<NumberingInstance>());
    }
}
