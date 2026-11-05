using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class OpenXmlExtensionsTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public OpenXmlExtensionsTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    [Theory]
    [InlineData(5)]
    [InlineData(200)]
    public void PreviousSiblingNearAgreesWithPreviousSiblingAsTheChildrenChange(int count)
    {
        Body body = new Body();
        for (int i = 0; i < count; i++)
        {
            body.AppendChild(new Paragraph());
        }

        void AssertAgrees(OpenXmlElement element)
        {
            Assert.Same(element.PreviousSibling(), element.PreviousSiblingNear());
        }

        //Forwards, backwards and jumping about, so that the element looked for is on either side of where the last one was.
        List<OpenXmlElement> children = body.ChildElements.ToList();
        children.ForEach(AssertAgrees);
        Enumerable.Reverse(children).ToList().ForEach(AssertAgrees);
        for (int i = 0; i < children.Count; i++)
        {
            AssertAgrees(children[i * 7 % children.Count]);
        }

        //The element last looked for and those around it are removed or have others put before them.
        OpenXmlElement last = children[children.Count - 1];
        AssertAgrees(last);
        last.PreviousSibling().Remove();
        AssertAgrees(last);
        body.InsertBefore(new Paragraph(), last);
        AssertAgrees(last);
        body.PrependChild(new Paragraph());
        AssertAgrees(last);
        OpenXmlElement middle = children[children.Count / 2];
        AssertAgrees(middle);
        foreach (OpenXmlElement child in children.Take(children.Count / 2))
        {
            child.Remove();
        }
        AssertAgrees(middle);
        AssertAgrees(last);
        body.ChildElements.ToList().ForEach(AssertAgrees);

        Assert.Null(body.FirstChild.PreviousSiblingNear());
        Assert.Null(new Paragraph().PreviousSiblingNear());
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void RootElementTableNestingLevelIsZero()
    {
        Assert.Equal(0, _document1.WordprocessingDocument.MainDocumentPart.Document.NestingLevel<Table>());
    }

    [Fact]
    public void SetNumAndSpaceColumnsSetsTheSpaceFromTheSpaceArgument()
    {
        Columns columns = new Columns();
        columns.SetNumAndSpaceColumns(2, 720);

        Assert.Equal(2, columns.ColumnCount.Value);
        Assert.Equal("720", columns.Space.Value);
    }

    [Fact]
    public void ConfigureColumnsAddsColumnsWhenTheSectionHasNone()
    {
        SectionProperties sectionProperties = new SectionProperties();
        sectionProperties.ConfigureColumns(3, 360);

        Assert.Equal(3, sectionProperties.GetColumns().ColumnCount.Value);
        Assert.Equal("360", sectionProperties.GetColumns().Space.Value);
    }

    [Fact]
    public void TruncateWithAZeroResolutionThrows()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => DateTime.UtcNow.Truncate(0));
    }

    [Fact]
    public void TruncateRemovesTheSubSecondPart()
    {
        DateTime date = new DateTime(2020, 1, 2, 3, 4, 5, 678, DateTimeKind.Utc);

        Assert.Equal(new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc), date.Truncate(TimeSpan.TicksPerSecond));
    }

    [Fact]
    public void GetTableWidthWithoutTablePropertiesIsZero()
    {
        Assert.Equal(0, new Table().GetTableWidth());
    }

    [Fact]
    public void GetTableWidthOfAnAutoWidthTableIsZero()
    {
        Table table = new Table(new TableProperties(new TableWidth() { Type = TableWidthUnitValues.Auto, Width = "0" }));

        Assert.Equal(0, table.GetTableWidth());
    }

    [Fact]
    public void GetTableWidthWithoutATypeDefaultsToDxa()
    {
        Table table = new Table(new TableProperties(new TableWidth() { Width = "5000" }));

        Assert.Equal(5000, table.GetTableWidth());
    }

    [Fact]
    public void GetTableWidthOfAPercentageTableUsesTheBodySectionPropertiesWhenTheTableIsNotLast()
    {
        Table table = new Table(new TableProperties(new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }));
        Body body = new Body(
            table,
            new Paragraph(),
            new SectionProperties(new PageSize() { Width = 12000, Height = 16000 }, new PageMargin() { Left = 1000, Right = 1000 }));

        //Page width less the page margins plus the default table cell margins.
        Assert.Equal(12000 - 2000 + 216, table.GetTableWidth());
    }

    [Fact]
    public void GetTableWidthOfAPercentageTableWithNoSectionPropertiesIsZero()
    {
        Table table = new Table(new TableProperties(new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }));
        Body body = new Body(table);

        Assert.Equal(0, table.GetTableWidth());
    }

    [Fact]
    public void SingleElementRangeIncludesTheElementItself()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a")));

        Assert.Single(paragraph.GetElements<Paragraph>(paragraph));
        Assert.Single(paragraph.GetElements<Run>(paragraph));
    }

    [Fact]
    public void GetBookmarkEndWorksOutsideTheMainDocument()
    {
        BookmarkStart bookmarkStart = new BookmarkStart() { Id = "1", Name = "a" };
        BookmarkEnd bookmarkEnd = new BookmarkEnd() { Id = "1" };
        Header header = new Header(new Paragraph(bookmarkStart, new Run(new Text("a")), bookmarkEnd));

        Assert.Same(bookmarkEnd, bookmarkStart.GetBookmarkEnd());
    }
}
