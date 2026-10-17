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
        Assert.Equal("PAGE \\* MERGEFORMAT", fields[0].Code);
        Assert.Equal("3", fields[0].Result);
        Assert.Equal(FieldType.NumPages, fields[1].Type);
        Assert.Equal("9", fields[1].Result);
        Assert.Equal(FieldType.Date, fields[2].Type);
        Assert.Equal("today", fields[2].Result);
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
        Assert.Equal("IF  = 1", fields[0].Code);
        Assert.Equal("first", fields[0].Result);
        Assert.Equal("PAGE", fields[1].Code);
        Assert.Equal("1", fields[1].Result);
    }

    [Fact]
    public void SettingTheResultKeepsTheFormattingOfTheFirstRun()
    {
        Paragraph paragraph = ComplexParagraph();
        FieldWrapper field = new RangeWrapper(paragraph).Fields[0];

        field.Result = "4";

        Assert.Equal("4", field.Result);
        Run resultRun = paragraph.Elements<Run>().Single(r => r.GetFirstChild<Text>()?.Text == "4");
        Assert.NotNull(resultRun.RunProperties.Bold);
        Assert.Equal("Page 4 of 9", new RangeWrapper(paragraph).Text.TrimEnd('\r'));
    }

    [Fact]
    public void SettingTheResultOfASimpleField()
    {
        FieldWrapper field = new RangeWrapper(CreateBody()).Fields[2];

        field.Result = "yesterday";

        Assert.Equal("yesterday", field.Result);
        Assert.NotNull(field.Range.StartElement.Descendants<Italic>().SingleOrDefault());
    }

    [Fact]
    public void SettingTheResultOfAFieldWithoutOneAddsTheSeparator()
    {
        Paragraph paragraph = new Paragraph(FieldCharRun(FieldCharValues.Begin), new Run(new FieldCode(" PAGE ")), FieldCharRun(FieldCharValues.End));
        FieldWrapper field = new RangeWrapper(paragraph).Fields.Single();
        Assert.Equal(string.Empty, field.Result);

        field.Result = "7";

        Assert.Equal("7", field.Result);
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
        Assert.Equal("PAGE", field.Code);

        field.Code = "NUMPAGES";

        Assert.Equal("NUMPAGES", field.Code);
        Assert.Equal(FieldType.NumPages, field.Type);
        Assert.Single(paragraph.Descendants<FieldCode>());
        Assert.Equal("1", field.Result);
    }

    [Fact]
    public void UnknownCodesAreOtherType()
    {
        Paragraph paragraph = new Paragraph(new SimpleField() { Instruction = " QUOTE \"x\" " });

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

    [Fact]
    public void AddAppendsADirtyComplexFieldToAParagraph()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("Page ")));
        RangeWrapper range = new RangeWrapper(paragraph);

        FieldWrapper field = range.Fields.Add(range, FieldType.Page, "\\* MERGEFORMAT");

        Assert.Equal(field, range.Fields.Single());
        Assert.Equal("PAGE \\* MERGEFORMAT", field.Code);
        Assert.Equal(FieldType.Page, field.Type);
        Assert.True(paragraph.Descendants<FieldChar>().First().Dirty.Value);
        Assert.Equal(3, paragraph.Descendants<FieldChar>().Count());

        field.Result = "1";
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
}
