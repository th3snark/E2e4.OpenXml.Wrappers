using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class BuiltInStylesTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public BuiltInStylesTest()
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
    public void EveryBuiltInStyleCanBeAskedForAndTheDocumentStaysValid()
    {
        //net48 has no generic Enum.GetValues.
#pragma warning disable CA2263
        foreach (BuiltInStyle style in Enum.GetValues(typeof(BuiltInStyle)))
#pragma warning restore CA2263
        {
            StyleWrapper wrapper = _document.Styles[style];
            Assert.NotNull(wrapper);
            Assert.True(wrapper.BuiltIn);
        }

        _document.Save();
        Assert.Empty(DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument));
    }

    [Fact]
    public void AStyleCanBeAppliedByNameToADocumentThatHasNoStyles()
    {
        MainDocumentPart main = _document.WordprocessingDocument.MainDocumentPart;
        main.DeletePart(main.StyleDefinitionsPart);
        _document.Content.Text = "Title text";

        _document.Paragraphs[0].Range.Style = "Heading 1";

        Assert.Equal("Heading 1", _document.Paragraphs[0].Range.Style);
        Assert.Equal("Heading1", _document.Paragraphs[0].Paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
        Assert.True(_document.Paragraphs[0].Range.Font.Bold == false);
        Assert.Equal(20, _document.Paragraphs[0].Range.Font.Size);
        Assert.True(_document.Paragraphs[0].Range.ParagraphFormat.KeepWithNext);
        Assert.Equal("Normal", _document.Styles["Heading 1"].BaseStyle);
    }

    [Fact]
    public void NamesAreFoundByTheNameWordShowsAndByTheStoredName()
    {
        StyleWrapper comment = _document.Styles["Comment Text"];

        Assert.Equal("Comment Text", comment.Name);
        Assert.Equal("annotation text", comment.Style.StyleName.Val.Value);
        Assert.Equal(comment, _document.Styles["annotation text"]);
        Assert.Equal(comment, _document.Styles["CommentText"]);
        Assert.Equal("TableNormal", _document.Styles["Table Grid"].Style.BasedOn.Val.Value);
    }

    [Fact]
    public void AnUnknownNameIsNotAStyle()
    {
        int count = _document.Styles.Count;

        Assert.Null(_document.Styles["No Such Style"]);
        Assert.Equal(count, _document.Styles.Count);
    }

    [Fact]
    public void AddingAParagraphStyleAddsItsLinkedStyleAfterIt()
    {
        StyleWrapper header = _document.Styles[BuiltInStyle.Header];

        Style[] styles = _document.WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().ToArray();
        Assert.Equal("HeaderChar", header.Style.LinkedStyle.Val.Value);
        Assert.Equal("HeaderChar", styles[Array.IndexOf(styles, header.Style) + 1].StyleId.Value);
    }

    [Fact]
    public void ApplyingAStyleUnhidesItAndItsLinkedStyleAsWordDoes()
    {
        _document.Content.Text = "Text";
        Style heading = _document.Styles["Heading 2"].Style;
        Style linked = _document.Styles["Heading 2 Char"].Style;
        Assert.NotNull(heading.SemiHidden);
        Assert.NotNull(linked.SemiHidden);

        _document.Paragraphs[0].Style = "Heading 2";

        Assert.Null(heading.SemiHidden);
        Assert.Null(linked.SemiHidden);
    }

    [Fact]
    public void AddingADefaultStyleKeepsTheOrderTheSchemaRequires()
    {
        Styles styles = _document.WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
        styles.Elements<Style>().Single(s => s.StyleId == "TableNormal").Remove();

        Assert.NotNull(_document.Styles[BuiltInStyle.NormalTable]);

        Assert.IsType<DocDefaults>(styles.ChildElements[0]);
        Assert.IsType<LatentStyles>(styles.ChildElements[1]);
        _document.Save();
        Assert.Empty(DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument));
    }

    [Fact]
    public void ACustomStyleKeepsItsName()
    {
        StyleWrapper custom = _document.Styles.Add("My Style", StyleType.Paragraph);

        Assert.Equal("My Style", custom.Name);
        Assert.False(custom.BuiltIn);
        custom.Delete();
        Assert.Null(_document.Styles["My Style"]);
    }
}
