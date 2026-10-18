using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class RangeWrapperStyleTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly Body _body;

    public RangeWrapperStyleTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        MainDocumentPart mainDocumentPart = _document.WordprocessingDocument.MainDocumentPart;
        StyleDefinitionsPart stylesPart = mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            new Style(new StyleName() { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
            new Style(new StyleName() { Val = "heading 1" }) { Type = StyleValues.Paragraph, StyleId = "Heading1" },
            new Style(new StyleName() { Val = "Strong" }) { Type = StyleValues.Character, StyleId = "Strong" },
            new Style(new StyleName() { Val = "Table Normal" }) { Type = StyleValues.Table, StyleId = "TableNormal", Default = true },
            new Style(new StyleName() { Val = "Table Grid" }) { Type = StyleValues.Table, StyleId = "TableGrid" },
            new Style(new StyleName() { Val = "No List" }) { Type = StyleValues.Numbering, StyleId = "NoList" });
        _body = mainDocumentPart.Document.Body;
        _body.RemoveAllChildren<Paragraph>();
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private Paragraph AddParagraph(string text = "a")
    {
        Paragraph paragraph = new Paragraph(new Run(new Text(text)));
        _body.InsertAt(paragraph, _body.Elements<Paragraph>().Count());
        return paragraph;
    }

    [Fact]
    public void AnUnstyledParagraphHasTheDefaultParagraphStyle()
    {
        Paragraph paragraph = AddParagraph();

        Assert.Equal("Normal", new RangeWrapper(paragraph).Style);
    }

    [Fact]
    public void SettingAParagraphStyleByNameIgnoresCaseAndReadsBackTheNameWordShows()
    {
        Paragraph paragraph = AddParagraph();
        RangeWrapper range = new RangeWrapper(paragraph);
        range.Style = "Heading 1";

        Assert.Equal("Heading1", paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
        Assert.Equal("Heading 1", range.Style);
    }

    [Fact]
    public void AStyleCanAlsoBeSetById()
    {
        Paragraph paragraph = AddParagraph();
        new RangeWrapper(paragraph).Style = "Heading1";

        Assert.Equal("Heading1", paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
    }

    [Fact]
    public void ANameBeatsAnIdThatIsTheNameOfAnotherStyle()
    {
        MainDocumentPart mainDocumentPart = _document.WordprocessingDocument.MainDocumentPart;
        Styles styles = mainDocumentPart.StyleDefinitionsPart.Styles;
        styles.AppendChild(new Style(new StyleName() { Val = "Heading1" }) { Type = StyleValues.Paragraph, StyleId = "Custom" });
        Paragraph paragraph = AddParagraph();

        new RangeWrapper(paragraph).Style = "Heading1";

        Assert.Equal("Custom", paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
    }

    [Fact]
    public void AnUnknownStyleIsRejectedAndLeavesTheParagraphUnchanged()
    {
        Paragraph paragraph = AddParagraph();

        Assert.Throws<ArgumentException>(() => new RangeWrapper(paragraph).Style = "No Such Style");
        Assert.Null(paragraph.ParagraphProperties);
    }

    [Fact]
    public void SettingAStyleOutsideADocumentWithStylesThrows()
    {
        Assert.Throws<InvalidOperationException>(() => new RangeWrapper(new Paragraph(new Run(new Text("a")))).Style = "Heading 1");
    }

    [Fact]
    public void ReadingTheStyleOutsideADocumentFallsBackToTheStyleId()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId() { Val = "Heading1" }));

        Assert.Equal("Heading1", new RangeWrapper(paragraph).Style);
    }

    [Fact]
    public void ThePargraphStyleIsWrittenFirstInTheParagraphProperties()
    {
        Paragraph paragraph = AddParagraph();
        paragraph.PrependChild(new ParagraphProperties(new KeepNext()));
        new RangeWrapper(paragraph).Style = "Heading 1";

        Assert.Equal(new[] { "pStyle", "keepNext" }, paragraph.ParagraphProperties.ChildElements.Select(e => e.LocalName).ToArray());
    }

    [Fact]
    public void TheStyleIsNullWhenTheParagraphsDiffer()
    {
        Paragraph first = AddParagraph();
        Paragraph second = AddParagraph();
        new RangeWrapper(first).Style = "Heading 1";

        Assert.Null(new RangeWrapper(_body).Style);

        new RangeWrapper(second).Style = "Heading 1";
        Assert.Equal("Heading 1", new RangeWrapper(_body).Style);
    }

    [Fact]
    public void SettingTheStyleOfARangeOfParagraphsStylesThemAll()
    {
        TableCell cell = new TableCell(new Paragraph(new Run(new Text("a"))), new Paragraph(new Run(new Text("b"))));
        Paragraph outside = AddParagraph();
        _body.InsertAt(new Table(new TableRow(cell)), 0);

        new RangeWrapper(cell).Style = "Heading 1";

        Assert.All(cell.Elements<Paragraph>(), p => Assert.Equal("Heading1", p.ParagraphProperties.ParagraphStyleId.Val.Value));
        Assert.Null(outside.ParagraphProperties);
    }

    [Fact]
    public void SettingTheStyleToNullRemovesIt()
    {
        Paragraph paragraph = AddParagraph();
        RangeWrapper range = new RangeWrapper(paragraph);
        range.Style = "Heading 1";
        range.Style = null;

        Assert.Empty(paragraph.ParagraphProperties.Elements<ParagraphStyleId>());
        Assert.Equal("Normal", range.Style);
    }

    [Fact]
    public void ACharacterStyleOnAParagraphRangeStylesItsRunsNotTheParagraph()
    {
        Paragraph paragraph = AddParagraph();
        RangeWrapper range = new RangeWrapper(paragraph);
        range.Style = "Strong";

        Assert.Equal("Strong", paragraph.Elements<Run>().Single().RunProperties.RunStyle.Val.Value);
        Assert.Null(paragraph.ParagraphProperties);
        Assert.Equal("Normal", range.Style);
    }

    [Fact]
    public void ABookmarkRangeReadsAndSetsTheCharacterStyleOfItsRuns()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Run outside = new Run(new Text("x"));
        Run inside = new Run(new Text("b"));
        Paragraph paragraph = new Paragraph(outside, start, inside, end);
        _body.InsertAt(paragraph, 0);
        RangeWrapper range = new RangeWrapper(start, end);

        Assert.Equal("Normal", range.Style);

        range.Style = "Strong";

        Assert.Equal("Strong", inside.RunProperties.RunStyle.Val.Value);
        Assert.Null(outside.RunProperties);
        Assert.Equal("Strong", range.Style);
    }

    [Fact]
    public void ABookmarkRangeAppliesAParagraphStyleToTheWholeParagraph()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Paragraph paragraph = new Paragraph(new Run(new Text("x")), start, new Run(new Text("b")), end);
        _body.InsertAt(paragraph, 0);

        new RangeWrapper(start, end).Style = "Heading 1";

        Assert.Equal("Heading1", paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
    }

    [Fact]
    public void ACharacterStyleSetToNullIsRemovedFromABookmarkRange()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Run inside = new Run(new Text("b"));
        _body.InsertAt(new Paragraph(start, inside, end), 0);
        RangeWrapper range = new RangeWrapper(start, end);
        range.Style = "Strong";
        range.Style = null;

        Assert.Null(inside.RunProperties);
    }

    [Fact]
    public void ATableStyleCannotBeAppliedToARange()
    {
        Paragraph paragraph = AddParagraph();

        Assert.Throws<NotSupportedException>(() => new RangeWrapper(paragraph).Style = "Table Grid");
    }

    [Fact]
    public void ANumberingStyleCannotBeApplied()
    {
        Paragraph paragraph = AddParagraph();

        Assert.Throws<NotSupportedException>(() => new RangeWrapper(paragraph).Style = "No List");
    }

    [Fact]
    public void ARangeThatHasNoTextHasNoStyle()
    {
        Assert.Throws<NotSupportedException>(() => new RangeWrapper(new Run(new Text("a"))).Style);
        Assert.Throws<NotSupportedException>(() => new RangeWrapper(new Run(new Text("a"))).Style = "Strong");
    }

    [Fact]
    public void TheStyleOfATableIsItsTableStyle()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        _body.InsertAt(table, 0);
        TableWrapper wrapper = new TableWrapper(table);
        Assert.Equal("Table Normal", wrapper.Style);

        wrapper.Style = "table grid";

        Assert.Equal("TableGrid", table.GetFirstChild<TableProperties>().TableStyle.Val.Value);
        Assert.IsType<TableProperties>(table.FirstChild);
        Assert.Equal("Table Grid", wrapper.Style);

        wrapper.Style = null;
        Assert.Null(table.GetFirstChild<TableProperties>().TableStyle);
        Assert.Equal("Table Normal", wrapper.Style);
    }

    [Fact]
    public void ATableStyleMustBeATableStyle()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        _body.InsertAt(table, 0);

        Assert.Throws<ArgumentException>(() => new TableWrapper(table).Style = "Heading 1");
        Assert.Throws<ArgumentException>(() => new TableWrapper(table).Style = "No Such Style");
        Assert.Null(table.GetFirstChild<TableProperties>());
    }

    [Fact]
    public void ATableStyleKeepsTheOtherTableProperties()
    {
        Table table = new Table(new TableProperties(new TableWidth() { Width = "100", Type = TableWidthUnitValues.Dxa }), new TableRow(new TableCell(new Paragraph())));
        _body.InsertAt(table, 0);
        new TableWrapper(table).Style = "Table Grid";

        Assert.Equal(new[] { "tblStyle", "tblW" }, table.GetFirstChild<TableProperties>().ChildElements.Select(e => e.LocalName).ToArray());
    }

    [Fact]
    public void StyleObjectReturnsTheStyleAsAnObject()
    {
        Paragraph paragraph = AddParagraph();
        RangeWrapper range = new RangeWrapper(paragraph);

        Assert.Equal(_document.Styles["Normal"], range.StyleObject);

        range.Style = "Heading 1";

        Assert.Equal(_document.Styles["Heading 1"], range.StyleObject);
        Assert.Equal("Heading 1", range.StyleObject.NameLocal);
        Assert.Equal(range.StyleObject, new ParagraphWrapper(paragraph).StyleObject);
    }

    [Fact]
    public void StyleObjectCanBeSetFromAStyleObjectAndNull()
    {
        Paragraph paragraph = AddParagraph();
        RangeWrapper range = new RangeWrapper(paragraph);

        range.StyleObject = _document.Styles["Heading 1"];
        Assert.Equal("Heading1", paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);

        range.StyleObject = null;
        Assert.Null(paragraph.ParagraphProperties.ParagraphStyleId);
    }

    [Fact]
    public void StyleObjectIsNullWhereTheStylesDiffer()
    {
        Paragraph first = AddParagraph();
        AddParagraph();
        new RangeWrapper(first).Style = "Heading 1";

        Assert.Null(new RangeWrapper(_body).StyleObject);
        Assert.Null(_document.Paragraphs.StyleObject);

        _document.Paragraphs.StyleObject = _document.Styles["Heading 1"];

        Assert.Equal(_document.Styles["Heading 1"], _document.Paragraphs.StyleObject);
    }

    [Fact]
    public void StyleObjectOfATableIsItsTableStyle()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        _body.InsertAt(table, 0);
        TableWrapper wrapper = new TableWrapper(table);

        wrapper.StyleObject = _document.Styles["Table Grid"];

        Assert.Equal("TableGrid", table.GetFirstChild<TableProperties>().TableStyle.Val.Value);
        Assert.Equal(_document.Styles["Table Grid"], wrapper.StyleObject);
        Assert.Throws<ArgumentException>(() => wrapper.StyleObject = _document.Styles["Heading 1"]);
    }

    [Fact]
    public void BaseAndNextStylesAreAvailableAsObjects()
    {
        StyleWrapper added = _document.Styles.Add("My Note", StyleType.Paragraph);

        Assert.Equal(_document.Styles["Normal"], added.BaseStyleObject);
        Assert.Null(_document.Styles["Normal"].BaseStyleObject);
        Assert.Null(added.NextParagraphStyleObject);
    }
}
