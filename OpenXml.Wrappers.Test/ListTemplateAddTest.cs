using System;
using System.Collections.Generic;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Adding list templates, the fonts and tab stops of their levels, and the members of lists. The markup is Word 16's (the list-template-*,
/// list-level-* and list-apply-* scenarios of MarkupConformanceTest); the values are those Word printed for the same calls.
/// </summary>
public class ListTemplateAddTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ListTemplateAddTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("One\rTwo\rThree\rFour\rFive");
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
    public void ASingleLevelTemplateHasTheDefaultsWordGives()
    {
        ListTemplateWrapper template = _document.ListTemplates.Add(false, "Mine");

        Assert.Equal("Mine", template.Name);
        Assert.False(template.OutlineNumbered);
        Assert.Equal(1, template.ListLevels.Count);
        ListLevelWrapper level = template.ListLevels[0];
        Assert.Equal("%1.", level.NumberFormat);
        Assert.Equal(ListNumberStyle.Arabic, level.NumberStyle);
        Assert.Equal(18, level.NumberPosition, 3);
        Assert.Equal(36, level.TextPosition, 3);
        Assert.Equal(9999999, level.TabPosition);
        Assert.Equal(TrailingCharacter.Tab, level.TrailingCharacter);
        Assert.Equal(1, level.StartAt);
        Assert.Equal(1, _document.ListTemplates.Count);
        Assert.Equal(0, _document.Lists.Count);
        AssertValid();
    }

    [Fact]
    public void AnOutlineTemplateHasNineLevelsAndAnUnnamedOneHasNoName()
    {
        ListTemplateWrapper outline = _document.ListTemplates.Add(true, "Outline");
        ListTemplateWrapper plain = _document.ListTemplates.Add();

        Assert.True(outline.OutlineNumbered);
        Assert.Equal(9, outline.ListLevels.Count);
        Assert.Equal("%2)", outline.ListLevels[1].NumberFormat);
        Assert.Equal(ListNumberStyle.LowercaseLetter, outline.ListLevels[1].NumberStyle);
        Assert.Equal("(%4)", outline.ListLevels[3].NumberFormat);
        Assert.Equal(string.Empty, plain.Name);
        Assert.Equal(2, _document.ListTemplates.Count);
        AssertValid();
    }

    [Fact]
    public void TheFontAndTabStopOfALevelAreSetAndRead()
    {
        ListTemplateWrapper template = _document.ListTemplates.Add(false, "Styled");
        ListLevelWrapper level = template.ListLevels[0];

        level.Font.Bold = true;
        level.Font.Size = 14;
        level.Font.Name = "Arial";
        level.Font.Color = 255;
        level.TabPosition = 100;

        Assert.True(level.Font.Bold);
        Assert.Equal(14, level.Font.Size);
        Assert.Equal("Arial", level.Font.Name);
        Assert.Equal(255, level.Font.Color.Value.Value);
        Assert.Equal(100, level.TabPosition, 3);

        level.Font.Bold = false;
        level.TabPosition = 9999999;
        Assert.False(level.Font.Bold);
        Assert.Equal(9999999, level.TabPosition);
        Assert.Throws<ArgumentOutOfRangeException>(() => level.TabPosition = -1);
        Format(0).ApplyListTemplate(template);
        AssertValid();
    }

    [Fact]
    public void ATemplateCanBeAppliedAtALevelOfItsOwn()
    {
        Format(0).ApplyNumberDefault();
        ListTemplateWrapper template = _document.ListTemplates[0];
        RangeWrapper range = _document.Range(_document.Paragraphs[1].Range.Start, _document.Paragraphs[2].Range.End);

        range.ListFormat.ApplyListTemplateWithLevel(template, true, ListApplyTo.WholeList, DefaultListBehavior.Word10, 2);

        Assert.Equal(1, Format(0).ListLevelNumber);
        Assert.Equal(2, Format(1).ListLevelNumber);
        Assert.Equal(2, Format(2).ListLevelNumber);
        Assert.Throws<ArgumentOutOfRangeException>(() => range.ListFormat.ApplyListTemplateWithLevel(template, false, ListApplyTo.WholeList, DefaultListBehavior.Word10, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => range.ListFormat.ApplyListTemplateWithLevel(template, false, ListApplyTo.WholeList, DefaultListBehavior.Word10, 10));
        Assert.Throws<NotSupportedException>(() => range.ListFormat.ApplyListTemplateWithLevel(template, false, ListApplyTo.ThisPointForward));
        AssertValid();
    }

    [Fact]
    public void ACopyOfATemplateInUseGetsTheNextFreeName()
    {
        ListTemplateWrapper template = _document.ListTemplates.Add(false, "A");
        Format(0).ApplyListTemplate(template);
        Format(1).ApplyListTemplate(template, false);
        Format(2).ApplyListTemplate(template, false);

        Assert.Equal(new[] { "A", "A22", "A2" }, _document.ListTemplates.Select(t => t.Name).ToArray());
        Assert.Equal(3, _document.Lists.Count);
        AssertValid();
    }

    [Fact]
    public void ANumberedParagraphBeforeTheRangeLetsTheListContinue()
    {
        ListTemplateWrapper a = _document.ListTemplates.Add(false, "A");
        ListTemplateWrapper b = _document.ListTemplates.Add(false, "B");

        Assert.Equal(ListContinue.Disabled, Format(0).CanContinuePreviousList(a));
        Format(0).ApplyListTemplate(a);
        Assert.Equal(ListContinue.ContinueList, Format(1).CanContinuePreviousList(a));
        Assert.Equal(ListContinue.ContinueList, Format(1).CanContinuePreviousList(b));
        Assert.Equal(ListContinue.ContinueList, Format(2).CanContinuePreviousList(a));
        Assert.Equal(ListContinue.Disabled, Format(0).CanContinuePreviousList(a));
        Assert.Throws<ArgumentNullException>(() => Format(1).CanContinuePreviousList(null));
    }

    [Fact]
    public void AListHasARangeAndCanBeChangedAsAWhole()
    {
        _document.Range(0, _document.Paragraphs[3].Range.End).ListFormat.ApplyNumberDefault();
        Assert.Equal(1, _document.Lists.Count);
        ListWrapper list = _document.Lists[0];

        Assert.Equal(0, list.Range.Start);
        Assert.Equal(_document.Paragraphs[3].Range.End, list.Range.End);

        ListTemplateWrapper template = _document.ListTemplates.Add(false, "Z");
        template.ListLevels[0].NumberFormat = "(%1)";
        list.ApplyListTemplate(template);
        Assert.Equal("(1)", Format(0).ListString);
        Assert.Equal("(4)", Format(3).ListString);
        Assert.Equal(ListContinue.Disabled, _document.Lists[0].CanContinuePreviousList(template));

        _document.Lists[0].RemoveNumbers();
        Assert.Equal(0, _document.Lists.Count);
    }

    [Fact]
    public void ConvertingAListToTextKeepsTheNumbersAsText()
    {
        Format(0).ApplyNumberDefault();
        Format(1).ApplyNumberDefault();

        _document.Lists[0].ConvertNumbersToText();

        Assert.Equal(0, _document.Lists.Count);
        Assert.StartsWith("1.\tOne\r2.\tTwo\r", _document.Content.Text, StringComparison.Ordinal);
    }
}
