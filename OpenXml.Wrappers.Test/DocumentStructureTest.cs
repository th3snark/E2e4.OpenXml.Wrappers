using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class DocumentStructureTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly Body _body;

    public DocumentStructureTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AddStyles()
    {
        StyleDefinitionsPart part = _document.WordprocessingDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        part.Styles = new Styles(
            new Style(new StyleName() { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
            new Style(new StyleName() { Val = "heading 1" }, new BasedOn() { Val = "Normal" }, new NextParagraphStyle() { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Heading1" },
            new Style(new StyleName() { Val = "Strong" }) { Type = StyleValues.Character, StyleId = "Strong" });
    }

    #region Sections

    [Fact]
    public void ANewDocumentHasOneSectionWithItsPageSetup()
    {
        Assert.Equal(1, _document.Sections.Count);
        PageSetupWrapper setup = _document.Sections[0].PageSetup;

        Assert.Equal(11906 / 20.0, setup.PageWidth);
        Assert.Equal(16838 / 20.0, setup.PageHeight);
        Assert.Equal(PageOrientation.Portrait, setup.Orientation);
        Assert.Equal(72, setup.TopMargin);
        Assert.Equal(90, setup.LeftMargin);
        Assert.Equal(90, setup.RightMargin);
        Assert.Equal(35.4, setup.HeaderDistance);
        Assert.Equal(0, setup.Gutter);
        Assert.Equal(SectionStart.NewPage, setup.SectionStart);
    }

    [Fact]
    public void SectionsAndTheirParentsLeadBackToTheDocument()
    {
        SectionWrapper section = _document.Sections[0];

        Assert.Equal(_document.Sections.Count, section.Parent.Count);
        Assert.Equal(section, section.Parent[0]);
        Assert.Equal(_document.Content, _document.Sections.Parent);
        Assert.Equal(section, section.Headers.Parent);
    }

    [Fact]
    public void SectionBreaksMakeMoreSectionsInOrder()
    {
        Paragraph first = new Paragraph(new ParagraphProperties(new SectionProperties(new PageSize() { Width = 1000, Height = 2000 })));
        Paragraph second = new Paragraph(new ParagraphProperties(new SectionProperties(new PageSize() { Width = 3000, Height = 4000 })));
        _body.InsertAt(first, 0);
        _body.InsertAt(new Paragraph(), 1);
        _body.InsertAt(second, 2);

        Assert.Equal(3, _document.Sections.Count);
        Assert.Equal(new[] { 50.0, 150.0, 11906 / 20.0 }, _document.Sections.Select(s => s.PageSetup.PageWidth).ToArray());
        Assert.Equal(new[] { 0, 1, 2 }, _document.Sections.Select(s => s.Index).ToArray());
        Assert.Same(first.ParagraphProperties.GetFirstChild<SectionProperties>(), _document.Sections[0].SectionProperties);
    }

    [Fact]
    public void ADocumentWithoutSectionPropertiesHasTheDefaultsUntilOneIsSet()
    {
        _body.RemoveAllChildren<SectionProperties>();
        PageSetupWrapper setup = _document.Sections[0].PageSetup;

        Assert.Equal(612, setup.PageWidth);
        Assert.Equal(792, setup.PageHeight);
        Assert.Equal(72, setup.LeftMargin);
        Assert.Equal(36, setup.FooterDistance);
        Assert.Empty(_body.Elements<SectionProperties>());

        setup.LeftMargin = 54;

        Assert.Equal(54, _document.Sections[0].PageSetup.LeftMargin);
        Assert.IsType<SectionProperties>(_body.LastChild);
    }

    [Fact]
    public void PageSetupValuesRoundTripInPoints()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;
        setup.PageWidth = 595.3;
        setup.PageHeight = 841.9;
        setup.TopMargin = 36;
        setup.BottomMargin = 40.5;
        setup.LeftMargin = 50;
        setup.RightMargin = 51;
        setup.Gutter = 10;
        setup.HeaderDistance = 20;
        setup.FooterDistance = 21;

        Assert.Equal(595.3, setup.PageWidth);
        Assert.Equal(841.9, setup.PageHeight);
        Assert.Equal(36, setup.TopMargin);
        Assert.Equal(40.5, setup.BottomMargin);
        Assert.Equal(50, setup.LeftMargin);
        Assert.Equal(51, setup.RightMargin);
        Assert.Equal(10, setup.Gutter);
        Assert.Equal(20, setup.HeaderDistance);
        Assert.Equal(21, setup.FooterDistance);
        Assert.Equal(720, _body.Elements<SectionProperties>().Single().GetFirstChild<PageMargin>().Top.Value);
    }

    [Fact]
    public void PageSetupRejectsBadValues()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;

        Assert.Throws<ArgumentOutOfRangeException>(() => setup.TopMargin = -1);
        Assert.Throws<ArgumentOutOfRangeException>(() => setup.PageWidth = 2000);
        Assert.Throws<ArgumentOutOfRangeException>(() => setup.LeftMargin = Double.NaN);
        Assert.Throws<ArgumentOutOfRangeException>(() => setup.Orientation = (PageOrientation)5);
        Assert.Throws<ArgumentOutOfRangeException>(() => setup.SectionStart = (SectionStart)9);
    }

    [Fact]
    public void ChangingTheOrientationSwapsTheSize()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;
        double width = setup.PageWidth;
        double height = setup.PageHeight;

        setup.Orientation = PageOrientation.Landscape;

        Assert.Equal(PageOrientation.Landscape, setup.Orientation);
        Assert.Equal(height, setup.PageWidth);
        Assert.Equal(width, setup.PageHeight);

        setup.Orientation = PageOrientation.Landscape;
        Assert.Equal(height, setup.PageWidth);

        setup.Orientation = PageOrientation.Portrait;
        Assert.Equal(width, setup.PageWidth);
    }

    [Fact]
    public void OrientationIsInferredFromTheSizeWhenNotStated()
    {
        _body.Elements<SectionProperties>().Single().GetFirstChild<PageSize>().Width = 20000;

        Assert.Equal(PageOrientation.Landscape, _document.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void SectionStartAndDifferentFirstPageRoundTrip()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;
        foreach (SectionStart start in EnumCompat.GetValues<SectionStart>())
        {
            setup.SectionStart = start;
            Assert.Equal(start, setup.SectionStart);
        }

        Assert.False(setup.DifferentFirstPageHeaderFooter);
        setup.DifferentFirstPageHeaderFooter = true;
        Assert.True(setup.DifferentFirstPageHeaderFooter);
        setup.DifferentFirstPageHeaderFooter = false;
        Assert.False(setup.DifferentFirstPageHeaderFooter);
        Assert.Empty(_body.Elements<SectionProperties>().Single().Elements<TitlePage>());
    }

    [Fact]
    public void SectionPropertiesStayInSchemaOrder()
    {
        _body.RemoveAllChildren<SectionProperties>();
        PageSetupWrapper setup = _document.Sections[0].PageSetup;
        setup.DifferentFirstPageHeaderFooter = true;
        setup.LeftMargin = 50;
        setup.PageWidth = 500;
        setup.SectionStart = SectionStart.Continuous;

        Assert.Equal(new[] { "type", "pgSz", "pgMar", "titlePg" }, _body.Elements<SectionProperties>().Single().ChildElements.Select(e => e.LocalName).ToArray());
    }

    [Fact]
    public void SectionsCompareByTheirProperties()
    {
        _body.InsertAt(new Paragraph(new ParagraphProperties(new SectionProperties())), 0);

        Assert.Equal(_document.Sections[0], _document.Sections[0]);
        Assert.NotEqual(_document.Sections[0], _document.Sections[1]);
        Assert.Equal(_document.Sections[1], _document.Sections[1]);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Sections[2]);
    }

    #endregion

    #region Styles

    [Fact]
    public void ADocumentWithoutStylesHasNone()
    {
        Assert.Equal(0, _document.Styles.Count);
        Assert.Null(_document.Styles["Normal"]);
        Assert.Empty(_document.Styles);
    }

    [Fact]
    public void StylesAreListedAndFoundByNameOrId()
    {
        AddStyles();

        Assert.Equal(3, _document.Styles.Count);
        Assert.Equal("heading 1", _document.Styles["Heading 1"].Name);
        Assert.Equal("Heading1", _document.Styles["Heading1"].StyleId);
        Assert.Equal("Strong", _document.Styles[2].Name);
        Assert.True(_document.Styles.Exists("normal"));
        Assert.False(_document.Styles.Exists("nothing"));
        Assert.Equal(new[] { "Normal", "heading 1", "Strong" }, _document.Styles.Select(s => s.Name).ToArray());
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Styles[3]);
    }

    [Fact]
    public void StyleDescribesItself()
    {
        AddStyles();
        StyleWrapper heading = _document.Styles["heading 1"];

        Assert.Equal(StyleType.Paragraph, heading.Type);
        Assert.Equal(StyleType.Character, _document.Styles["Strong"].Type);
        Assert.True(heading.BuiltIn);
        Assert.Equal("Normal", heading.BaseStyle);
        Assert.Equal("Normal", heading.NextParagraphStyle);
        Assert.Null(_document.Styles["Normal"].BaseStyle);
        Assert.Equal(heading, _document.Styles["Heading1"]);
    }

    [Fact]
    public void AStyleCanBeAddedAndUsed()
    {
        AddStyles();
        StyleWrapper added = _document.Styles.Add("My Note", StyleType.Paragraph);

        Assert.Equal("My Note", added.Name);
        Assert.Equal("MyNote", added.StyleId);
        Assert.False(added.BuiltIn);
        Assert.Equal("Normal", added.BaseStyle);
        Assert.Equal(4, _document.Styles.Count);

        Paragraph paragraph = new Paragraph(new Run(new Text("x")));
        _body.InsertAt(paragraph, 0);
        new RangeWrapper(paragraph).Style = "my note";
        Assert.Equal("MyNote", paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
    }

    [Fact]
    public void AddingToADocumentWithoutStylesCreatesThePart()
    {
        StyleWrapper added = _document.Styles.Add("First", StyleType.Character);

        Assert.Equal(StyleType.Character, added.Type);
        Assert.NotNull(_document.WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart);
        Assert.Equal(1, _document.Styles.Count);
    }

    [Theory]
    [InlineData("Normal")]
    [InlineData("normal")]
    [InlineData("Heading1")]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData("!!!")]
    public void AddingABadOrDuplicateStyleIsRejected(string name)
    {
        AddStyles();

        Assert.Throws<ArgumentException>(() => _document.Styles.Add(name, StyleType.Paragraph));
        Assert.Equal(3, _document.Styles.Count);
    }

    [Fact]
    public void DeletingAStyleResetsItsUsers()
    {
        AddStyles();
        StyleWrapper custom = _document.Styles.Add("Custom", StyleType.Paragraph);
        StyleWrapper child = _document.Styles.Add("Child", StyleType.Paragraph);
        child.Style.GetFirstChild<BasedOn>().Val = "Custom";
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId() { Val = "Custom" }), new Run(new RunProperties(new RunStyle() { Val = "Custom" }), new Text("x")));
        _body.InsertAt(paragraph, 0);

        _document.Styles.Delete(custom);

        Assert.False(_document.Styles.Exists("Custom"));
        Assert.Null(child.BaseStyle);
        Assert.Null(paragraph.ParagraphProperties.ParagraphStyleId);
        Assert.Null(paragraph.Elements<Run>().Single().RunProperties.RunStyle);
    }

    [Fact]
    public void BuiltInAndDefaultStylesCannotBeDeleted()
    {
        AddStyles();

        Assert.Throws<InvalidOperationException>(() => _document.Styles.Delete(_document.Styles["Normal"]));
        Assert.Throws<InvalidOperationException>(() => _document.Styles.Delete(_document.Styles["heading 1"]));
        Assert.Equal(3, _document.Styles.Count);
    }

    #endregion

    #region Properties

    [Fact]
    public void DocumentPropertiesRoundTrip()
    {
        DocumentPropertiesWrapper properties = _document.BuiltInDocumentProperties;
        Assert.Null(properties.Title);

        properties.Title = "Report";
        properties.Subject = "Q3";
        properties.Author = "Pat";
        properties.Keywords = "a, b";
        properties.Comments = "Draft";
        properties.Category = "Finance";
        properties.LastAuthor = "Sam";
        DateTime created = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        properties.CreationDate = created;
        properties.LastSaveTime = created.AddDays(1);

        Assert.Equal("Report", properties.Title);
        Assert.Equal("Q3", properties.Subject);
        Assert.Equal("Pat", properties.Author);
        Assert.Equal("a, b", properties.Keywords);
        Assert.Equal("Draft", properties.Comments);
        Assert.Equal("Finance", properties.Category);
        Assert.Equal("Sam", properties.LastAuthor);
        Assert.Equal(created, properties.CreationDate);
        Assert.Equal(created.AddDays(1), properties.LastSaveTime);

        properties.Title = null;
        Assert.Null(properties.Title);
    }

    [Fact]
    public void ExtendedPropertiesAreCreatedOnDemand()
    {
        DocumentPropertiesWrapper properties = _document.BuiltInDocumentProperties;
        Assert.Null(properties.Company);
        properties.Company = null;
        Assert.Null(_document.WordprocessingDocument.ExtendedFilePropertiesPart);

        properties.Company = "Acme";
        properties.Manager = "Lee";
        Assert.Equal("Acme", properties.Company);
        Assert.Equal("Lee", properties.Manager);

        properties.Company = null;
        Assert.Null(properties.Company);
        Assert.Equal("Lee", properties.Manager);
    }

    [Fact]
    public void PropertiesSurviveASaveAndReload()
    {
        _document.BuiltInDocumentProperties.Title = "Kept";
        _document.BuiltInDocumentProperties.Company = "Acme";
        string copy = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            _document.SaveAs(copy).Close();
            using WordprocessingDocumentWrapper reopened = WordprocessingDocumentWrapper.Open(copy, false);
            Assert.Equal("Kept", reopened.BuiltInDocumentProperties.Title);
            Assert.Equal("Acme", reopened.BuiltInDocumentProperties.Company);
        }
        finally
        {
            File.Delete(copy);
        }
    }

    #endregion
}
