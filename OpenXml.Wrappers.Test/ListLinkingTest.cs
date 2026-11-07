using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Linking list levels to styles, converting list templates, and the look of tables of contents. The expected values are those Word 16 gave
/// for the same calls (the list-linked-style, list-linked-style-applied, list-convert-* and toc-format-* scenarios of
/// tools/WordReference/scenarios.ps1).
/// </summary>
public class ListLinkingTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ListLinkingTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("One\rTwo");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ALevelCanBeLinkedToAStyleAndFreedAgain()
    {
        ListTemplateWrapper template = _document.ListTemplates.Add(true, "Linked");
        Assert.Equal(string.Empty, template.ListLevels[0].LinkedStyle);
        template.ListLevels[0].LinkedStyle = "Heading 1";
        template.ListLevels[1].LinkedStyle = "Heading 2";
        Assert.Equal("Heading 1", template.ListLevels[0].LinkedStyle);
        Assert.Equal("Heading 2", template.ListLevels[1].LinkedStyle);

        //A style is linked to one level only.
        template.ListLevels[2].LinkedStyle = "Heading 1";
        Assert.Equal(string.Empty, template.ListLevels[0].LinkedStyle);
        Assert.Equal("Heading 1", template.ListLevels[2].LinkedStyle);

        template.ListLevels[2].LinkedStyle = string.Empty;
        Assert.Equal(string.Empty, template.ListLevels[2].LinkedStyle);
        Assert.Throws<ArgumentException>(() => template.ListLevels[0].LinkedStyle = "No such style");
        Assert.Throws<ArgumentNullException>(() => template.ListLevels[0].LinkedStyle = null);
    }

    [Fact]
    public void ApplyingALinkedTemplateGivesTheParagraphsTheStyle()
    {
        ListTemplateWrapper template = _document.ListTemplates.Add(true, "Linked");
        template.ListLevels[0].LinkedStyle = "Heading 1";
        _document.Paragraphs[0].Range.ListFormat.ApplyListTemplate(template);
        Assert.Equal("Heading 1", _document.Paragraphs[0].Range.ParagraphStyle);
        Assert.Single(_document.ListTemplates);
        Assert.Single(_document.Lists);
    }

    [Fact]
    public void ConvertingTemplates()
    {
        ListTemplateWrapper outline = _document.ListTemplates.Add(true, "Outline");
        outline.Convert(3);
        Assert.False(outline.OutlineNumbered);
        Assert.Equal(1, outline.ListLevels.Count);
        Assert.Equal("%3)", outline.ListLevels[0].NumberFormat);

        ListTemplateWrapper single = _document.ListTemplates.Add(false, "Single");
        single.Convert();
        Assert.True(single.OutlineNumbered);
        Assert.Equal(9, single.ListLevels.Count);
        Assert.Equal("%1.", single.ListLevels[0].NumberFormat);
        Assert.Equal("%2)", single.ListLevels[1].NumberFormat);

        ListTemplateWrapper fresh = _document.ListTemplates.Add(false);
        Assert.Throws<ArgumentOutOfRangeException>(() => fresh.Convert(2));
        ListTemplateWrapper another = _document.ListTemplates.Add(true);
        Assert.Throws<ArgumentOutOfRangeException>(() => another.Convert(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => another.Convert(10));
    }

    [Fact]
    public void TheLookOfTablesOfContents()
    {
        _document.Content.InsertAfter("\rSecond");
        _document.Paragraphs[0].Style = "Heading 1";
        _document.Paragraphs[2].Style = "Heading 1";
        _document.TablesOfContents.Add(_document.Range(0, 0));
        Assert.Equal(TocFormat.Template, _document.TablesOfContents.Format);

        foreach (TocFormat format in new[] { TocFormat.Classic, TocFormat.Distinctive, TocFormat.Fancy, TocFormat.Modern, TocFormat.Formal, TocFormat.Simple })
        {
            _document.TablesOfContents.Format = format;
            Assert.Equal(format, _document.TablesOfContents.Format);
        }

        //The template look changes nothing.
        _document.TablesOfContents.Format = TocFormat.Fancy;
        _document.TablesOfContents.Format = TocFormat.Template;
        Assert.Equal(TocFormat.Fancy, _document.TablesOfContents.Format);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.TablesOfContents.Format = (TocFormat)9);
        Assert.Equal(9, new[] { "TOC 1", "TOC 2", "TOC 3", "TOC 4", "TOC 5", "TOC 6", "TOC 7", "TOC 8", "TOC 9" }.Count(name => _document.Styles[name] != null));
    }

    [Fact]
    public void TheMarkupIsValid()
    {
        _document.Content.InsertAfter("\rSecond");
        ListTemplateWrapper template = _document.ListTemplates.Add(true, "Linked");
        template.ListLevels[0].LinkedStyle = "Heading 1";
        template.ListLevels[1].LinkedStyle = "Heading 2";
        _document.Paragraphs[0].Range.ListFormat.ApplyListTemplate(template);
        _document.ListTemplates.Add(true, "Outline").Convert(3);
        _document.ListTemplates.Add(false, "Single").Convert();
        _document.TablesOfContents.Add(_document.Range(0, 0));
        _document.TablesOfContents.Format = TocFormat.Modern;
        _document.Save();
        System.Collections.Generic.IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }
}
