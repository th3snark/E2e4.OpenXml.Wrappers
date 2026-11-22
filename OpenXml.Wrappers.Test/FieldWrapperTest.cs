using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class FieldWrapperTest
{
    private static Run FieldCharRun(FieldCharValues type)
    {
        return new Run(new FieldChar() { FieldCharType = type });
    }

    private static Paragraph ComplexParagraph()
    {
        //Page " of " NumPages, the first with a bold result.
        return new Paragraph(
            new Run(new Text("Page ")),
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" PAGE \\* MERGEFORMAT ")),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new RunProperties(new Bold()), new Text("3")),
            FieldCharRun(FieldCharValues.End),
            new Run(new Text(" of ")),
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" NUMPAGES ")),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new Text("9")),
            FieldCharRun(FieldCharValues.End));
    }

    private static Body CreateBody()
    {
        return new Body(
            ComplexParagraph(),
            new Paragraph(new SimpleField(new Run(new RunProperties(new Italic()), new Text("today"))) { Instruction = " DATE \\@ \"d MMMM yyyy\" " }));
    }

    [Fact]
    public void FindsComplexAndSimpleFieldsInOrder()
    {
        FieldWrapperCollection fields = new RangeWrapper(CreateBody()).Fields;

        Assert.Equal(3, fields.Count);
        Assert.Equal(FieldType.Page, fields[0].Type);
        Assert.Equal("PAGE \\* MERGEFORMAT", fields[0].Code.Text);
        Assert.Equal("3", fields[0].Result.Text);
        Assert.Equal(FieldType.NumPages, fields[1].Type);
        Assert.Equal("9", fields[1].Result.Text);
        Assert.Equal(FieldType.Date, fields[2].Type);
        Assert.Equal("today", fields[2].Result.Text);
        Assert.Equal(3, fields.ToList().Count);
        Assert.Throws<System.ArgumentOutOfRangeException>(() => fields[3]);
    }

    [Fact]
    public void RangeOfOneParagraphOnlyHasItsFields()
    {
        Body body = CreateBody();

        Assert.Equal(2, new RangeWrapper(body.Elements<Paragraph>().First()).Fields.Count);
        Assert.Single(new RangeWrapper(body.Elements<Paragraph>().Last()).Fields);
    }

    [Fact]
    public void NestedFieldsAreListedAfterTheirOuterField()
    {
        Paragraph paragraph = new Paragraph(
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" IF ")),
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" PAGE ")),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new Text("1")),
            FieldCharRun(FieldCharValues.End),
            new Run(new FieldCode(" = 1 ")),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new Text("first")),
            FieldCharRun(FieldCharValues.End));
        FieldWrapperCollection fields = new RangeWrapper(paragraph).Fields;

        Assert.Equal(2, fields.Count);
        Assert.Equal("IF  = 1", fields[0].Code.Text);
        Assert.Equal("first", fields[0].Result.Text);
        Assert.Equal("PAGE", fields[1].Code.Text);
        Assert.Equal("1", fields[1].Result.Text);
    }

    [Fact]
    public void SettingTheResultKeepsTheFormattingOfTheFirstRun()
    {
        Paragraph paragraph = ComplexParagraph();
        FieldWrapper field = new RangeWrapper(paragraph).Fields[0];

        field.Result.Text = "4";

        Assert.Equal("4", field.Result.Text);
        Run resultRun = paragraph.Elements<Run>().Single(r => r.GetFirstChild<Text>()?.Text == "4");
        Assert.NotNull(resultRun.RunProperties.Bold);
        Assert.Equal("Page 4 of 9", new RangeWrapper(paragraph).Text.TrimEnd('\r'));
    }

    [Fact]
    public void SettingTheResultOfASimpleField()
    {
        FieldWrapper field = new RangeWrapper(CreateBody()).Fields[2];

        field.Result.Text = "yesterday";

        Assert.Equal("yesterday", field.Result.Text);
        Assert.NotNull(field.Range.StartElement.Descendants<Italic>().SingleOrDefault());
    }

    [Fact]
    public void SettingTheResultOfAFieldWithoutOneAddsTheSeparator()
    {
        Paragraph paragraph = new Paragraph(FieldCharRun(FieldCharValues.Begin), new Run(new FieldCode(" PAGE ")), FieldCharRun(FieldCharValues.End));
        FieldWrapper field = new RangeWrapper(paragraph).Fields.Single();
        Assert.Equal(string.Empty, field.Result.Text);

        field.Result.Text = "7";

        Assert.Equal("7", field.Result.Text);
        Assert.Single(paragraph.Descendants<FieldChar>(), f => f.FieldCharType.Value == FieldCharValues.Separate);
    }

    [Fact]
    public void SettingTheCodeReplacesAllCodeRuns()
    {
        Paragraph paragraph = new Paragraph(
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" PA")),
            new Run(new FieldCode("GE ")),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new Text("1")),
            FieldCharRun(FieldCharValues.End));
        FieldWrapper field = new RangeWrapper(paragraph).Fields.Single();
        Assert.Equal("PAGE", field.Code.Text);

        field.Code.Text = "NUMPAGES";

        Assert.Equal("NUMPAGES", field.Code.Text);
        Assert.Equal(FieldType.NumPages, field.Type);
        Assert.Single(paragraph.Descendants<FieldCode>());
        Assert.Equal("1", field.Result.Text);
    }

    [Fact]
    public void UnknownCodesAreOtherType()
    {
        Paragraph paragraph = new Paragraph(new SimpleField() { Instruction = " NOSUCHFIELD \"x\" " });

        Assert.Equal(FieldType.Other, new RangeWrapper(paragraph).Fields.Single().Type);
    }

    [Fact]
    public void DeleteRemovesTheWholeField()
    {
        Paragraph paragraph = ComplexParagraph();
        RangeWrapper range = new RangeWrapper(paragraph);

        range.Fields[0].Delete();

        Assert.Single(range.Fields);
        Assert.Equal("Page  of 9", range.Text.TrimEnd('\r'));
    }

    //"Before X after." with a PAGE field at 7, as Word lays it out: begin 7, code 8-14, separator 14, result 15, end 16. The counts are what
    //Word gives for Document.Range(start, end).Fields.Count: a span has a field when it covers the field's begin or end.
    [Theory]
    [InlineData(0, 3, 0)]
    [InlineData(0, 7, 0)]
    [InlineData(0, 8, 1)]
    [InlineData(0, 17, 1)]
    [InlineData(5, 9, 1)]
    [InlineData(6, 7, 0)]
    [InlineData(7, 7, 0)]
    [InlineData(7, 8, 1)]
    [InlineData(8, 8, 0)]
    [InlineData(10, 12, 0)]
    [InlineData(14, 14, 0)]
    [InlineData(15, 16, 0)]
    [InlineData(16, 16, 0)]
    [InlineData(16, 17, 1)]
    [InlineData(17, 20, 0)]
    public void FieldsOfASpanAreThoseWhoseBeginOrEndItCovers(int start, int end, int expected)
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Paragraphs[0].Range.Text = "Before X after.";
        FieldWrapper field = document.Fields.Add(document.Range(7, 7), FieldType.Empty, "PAGE", false);
        field.Result.Text = "1";
        Assert.Equal(8, field.Code.Range.Start);
        Assert.Equal(14, field.Code.Range.End);
        Assert.Equal(15, field.Result.Range.Start);
        Assert.Equal(16, field.Result.Range.End);

        FieldWrapperCollection fields = document.Range(start, end).Fields;

        Assert.Equal(expected, fields.Count);
        Assert.Equal(expected, fields.ToList().Count);
        if (expected == 1)
        {
            Assert.Equal(field, fields[0]);
        }
    }

    [Fact]
    public void FieldsOfASpanIncludeSimpleAndNestedFieldsInOrder()
    {
        Body body = CreateBody();
        Paragraph nested = new Paragraph(
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" IF ")),
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" PAGE ")),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new Text("1")),
            FieldCharRun(FieldCharValues.End),
            new Run(new FieldCode(" = 1 yes no ")),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new Text("yes")),
            FieldCharRun(FieldCharValues.End));
        body.AppendChild(nested);
        StoryMap map = StoryMap.Build(body);
        List<FieldWrapper> all = new RangeWrapper(body).Fields.ToList();
        Assert.Equal(5, all.Count);

        //The whole story.
        Assert.Equal(all, RangeWrapper.FromMap(map, 0, map.Length).Fields.ToList());

        //The date field, written as one element, from its begin or its end.
        FieldWrapper date = all[2];
        SimpleField simple = body.Descendants<SimpleField>().Single();
        Atom begin = map.Atoms.Single(a => ReferenceEquals(a.Element, simple));
        Atom end = map.Atoms.Single(a => a.Kind == AtomKind.FieldMark && ReferenceEquals(a.Element.Parent, simple));
        Assert.Equal(new[] { date }, RangeWrapper.FromMap(map, begin.Start, begin.Start + 1).Fields.ToList());
        Assert.Equal(new[] { date }, RangeWrapper.FromMap(map, end.Start, end.End).Fields.ToList());
        Assert.Empty(RangeWrapper.FromMap(map, begin.End, end.Start).Fields);

        //From the end of the inner field to the end of the outer one: both, outer first.
        FieldWrapper outer = all[3];
        FieldWrapper inner = all[4];
        Assert.Equal(new[] { outer, inner }, RangeWrapper.FromMap(map, inner.Result.Range.End, outer.Result.Range.End + 1).Fields.ToList());
    }

    [Fact]
    public void FieldsOfTheRangeOfANote()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        ParagraphWrapper paragraph = document.Paragraphs[0];
        paragraph.Range.Text = "A claim.";
        FootnoteWrapper footnote = paragraph.Range.Footnotes.Add(paragraph.Range, null, "See page ");
        EndnoteWrapper endnote = paragraph.Range.Endnotes.Add(paragraph.Range, null, "See page ");
        Assert.Empty(footnote.Range.Fields);
        Assert.Empty(endnote.Range.Fields);

        RangeWrapper footnoteText = footnote.Range;
        FieldWrapper inFootnote = footnoteText.Fields.Add(footnoteText, "PAGE");
        inFootnote.Result.Text = "1";
        RangeWrapper endnoteText = endnote.Range;
        FieldWrapper inEndnote = endnoteText.Fields.Add(endnoteText, "PAGE");
        inEndnote.Result.Text = "1";

        Assert.Equal(inFootnote, footnote.Range.Fields.Single());
        Assert.Equal(inEndnote, endnote.Range.Fields.Single());
        Assert.Equal("See page 1", footnote.Range.Text);
    }

    [Fact]
    public void AddAppendsADirtyComplexFieldToAParagraph()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("Page ")));
        RangeWrapper range = new RangeWrapper(paragraph);

        FieldWrapper field = range.Fields.Add(range, FieldType.Page, "\\* MERGEFORMAT");

        Assert.Equal(field, range.Fields.Single());
        Assert.Equal("PAGE \\* MERGEFORMAT", field.Code.Text);
        Assert.Equal(FieldType.Page, field.Type);
        Assert.True(paragraph.Descendants<FieldChar>().First().Dirty.Value);
        Assert.Equal(3, paragraph.Descendants<FieldChar>().Count());

        field.Result.Text = "1";
        Assert.Equal("Page 1", range.Text.TrimEnd('\r'));
    }

    [Fact]
    public void AddWithACodeAndToABody()
    {
        Body body = new Body();
        RangeWrapper range = new RangeWrapper(body);

        FieldWrapper field = range.Fields.Add(range, "MERGEFIELD Name");

        Assert.Equal(FieldType.MergeField, field.Type);
        Assert.Single(body.Elements<Paragraph>());
        Assert.Throws<System.ArgumentException>(() => range.Fields.Add(range, " "));
        Assert.Throws<System.ArgumentOutOfRangeException>(() => range.Fields.Add(range, FieldType.Other));
    }

    [Fact]
    public void ParentRangeAndEquality()
    {
        Body body = CreateBody();
        FieldWrapper field = new RangeWrapper(body).Fields[1];

        Assert.Equal(new RangeWrapper(body.Elements<Paragraph>().First()), field.Parent);
        Assert.Equal(field, new RangeWrapper(body).Fields[1]);
        Assert.NotEqual(field, new RangeWrapper(body).Fields[0]);
    }

    [Fact]
    public void ResultRangeCoversTheResultCharacters()
    {
        Body body = CreateBody();
        FieldWrapper page = new RangeWrapper(body).Fields[0];

        RangeWrapper result = page.Result.Range;

        Assert.Equal("3", result.Text);
        //Like any range, the code range does not show field codes in its Text (Field.Code.Text has them), but it covers their characters.
        Assert.Equal(" PAGE \\* MERGEFORMAT ".Length, page.Code.Range.End - page.Code.Range.Start);
    }

    [Fact]
    public void ResultFontFormatsOnlyTheResult()
    {
        Body body = CreateBody();
        FieldWrapper numPages = new RangeWrapper(body).Fields[1];

        numPages.Result.Font.Italic = true;

        Run resultRun = body.Elements<Paragraph>().First().Elements<Run>().Last(r => r.Elements<Text>().Any(t => t.Text == "9"));
        Assert.NotNull(resultRun.RunProperties.Italic);
        Assert.Equal(true, numPages.Result.Font.Italic);
        Assert.Equal(true, new RangeWrapper(body).Fields[1].Result.Font.Italic);
        Assert.NotEqual(true, new RangeWrapper(body).Fields[0].Result.Font.Italic);
        Assert.Equal(true, new RangeWrapper(body).Fields[0].Result.Font.Bold);
    }

    [Fact]
    public void ResultOfASimpleFieldIsItsRuns()
    {
        Body body = CreateBody();
        FieldWrapper date = new RangeWrapper(body).Fields[2];

        Assert.Equal("today", date.Result.Range.Text);
        Assert.Equal(true, date.Result.Font.Italic);
        Assert.Equal(0, date.Code.Range.End - date.Code.Range.Start);
    }

    [Fact]
    public void ResultRangeOfAFieldWithoutResultIsEmpty()
    {
        Body body = new Body(new Paragraph(
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" PAGE ")),
            FieldCharRun(FieldCharValues.End)));
        FieldWrapper field = new RangeWrapper(body).Fields[0];

        RangeWrapper result = field.Result.Range;

        Assert.Equal(result.Start, result.End);
    }
}
