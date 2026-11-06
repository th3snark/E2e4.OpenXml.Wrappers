using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Reading lists: ListTemplates, ListLevels, Lists, ListParagraphs and ListFormat. The numbering definitions are the ones Word 16 writes for
/// its default bullet, number and outline lists, and the values are what Word answered for them.
/// </summary>
public class ListFormatReadTest : IDisposable
{
    private const string Bullet = "";
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ListFormatReadTest()
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

    private static Level NewLevel(int index, NumberFormatValues format, string text, int left, int hanging)
    {
        return new Level(
            new StartNumberingValue() { Val = 1 },
            new NumberingFormat() { Val = format },
            new LevelText() { Val = text },
            new LevelJustification() { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(new Indentation() { Left = left.ToString(System.Globalization.CultureInfo.InvariantCulture), Hanging = hanging.ToString(System.Globalization.CultureInfo.InvariantCulture) }))
        { LevelIndex = index };
    }

    private void AddNumbering(params AbstractNum[] templates)
    {
        MainDocumentPart main = _document.WordprocessingDocument.MainDocumentPart;
        NumberingDefinitionsPart part = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = new Numbering();
        int number = 1;
        foreach (AbstractNum template in templates)
        {
            part.Numbering.AppendChild(template);
        }
        foreach (AbstractNum template in templates)
        {
            part.Numbering.AppendChild(new NumberingInstance(new AbstractNumId() { Val = template.AbstractNumberId.Value }) { NumberID = number++ });
        }
    }

    private static AbstractNum BulletTemplate(int id)
    {
        return new AbstractNum(new MultiLevelType() { Val = MultiLevelValues.SingleLevel }, NewLevel(0, NumberFormatValues.Bullet, Bullet, 720, 360)) { AbstractNumberId = id };
    }

    private static AbstractNum NumberTemplate(int id)
    {
        return new AbstractNum(new MultiLevelType() { Val = MultiLevelValues.SingleLevel }, NewLevel(0, NumberFormatValues.Decimal, "%1.", 720, 360)) { AbstractNumberId = id };
    }

    private static AbstractNum OutlineTemplate(int id)
    {
        return new AbstractNum(
            new MultiLevelType() { Val = MultiLevelValues.Multilevel },
            NewLevel(0, NumberFormatValues.Decimal, "%1)", 360, 360),
            NewLevel(1, NumberFormatValues.LowerLetter, "%2)", 720, 360),
            NewLevel(2, NumberFormatValues.LowerRoman, "%3)", 1080, 360),
            NewLevel(3, NumberFormatValues.Decimal, "(%4)", 1440, 360),
            NewLevel(4, NumberFormatValues.LowerLetter, "(%5)", 1800, 360))
        { AbstractNumberId = id };
    }

    private void Number(int paragraph, int numId, int level = 0)
    {
        Paragraph element = _document.Paragraphs[paragraph].Paragraph;
        element.GetOrAddParagraphProperties().NumberingProperties = new NumberingProperties(new NumberingLevelReference() { Val = level }, new NumberingId() { Val = numId });
    }

    [Fact]
    public void AParagraphInNoListHasNoListFormat()
    {
        ListFormatWrapper format = _document.Paragraphs[0].Range.ListFormat;

        Assert.Equal(ListType.NoNumbering, format.ListType);
        Assert.Equal(1, format.ListLevelNumber);
        Assert.Equal(string.Empty, format.ListString);
        Assert.Equal(0, format.ListValue);
        Assert.Null(format.ListTemplate);
        Assert.Null(format.List);
        Assert.Equal(0, format.CountNumberedItems());
        Assert.Equal(0, _document.Lists.Count);
        Assert.Equal(0, _document.ListParagraphs.Count);
        Assert.Equal(0, _document.ListTemplates.Count);
    }

    [Fact]
    public void ABulletedListReadsAsWordReadsIt()
    {
        AddNumbering(BulletTemplate(0));
        for (int i = 0; i < 3; i++)
        {
            Number(i, 1);
        }

        ListFormatWrapper second = _document.Paragraphs[1].Range.ListFormat;

        Assert.Equal(ListType.Bullet, second.ListType);
        Assert.Equal(1, second.ListLevelNumber);
        Assert.Equal(Bullet, second.ListString);
        Assert.Equal(2, second.ListValue);
        Assert.Equal(3, _document.Content.ListFormat.CountNumberedItems());
        Assert.Equal(1, _document.Lists.Count);
        Assert.Equal(3, _document.ListParagraphs.Count);
        Assert.Equal(1, _document.ListTemplates.Count);
    }

    [Fact]
    public void ANumberedListCountsFromOne()
    {
        AddNumbering(NumberTemplate(0));
        for (int i = 0; i < 3; i++)
        {
            Number(i, 1);
        }

        Assert.Equal(ListType.SimpleNumbering, _document.Paragraphs[1].Range.ListFormat.ListType);
        Assert.Equal("1.", _document.Paragraphs[0].Range.ListFormat.ListString);
        Assert.Equal("2.", _document.Paragraphs[1].Range.ListFormat.ListString);
        Assert.Equal("3.", _document.Paragraphs[2].Range.ListFormat.ListString);
        Assert.Equal(ListType.MixedNumbering, _document.Content.ListFormat.ListType);
        Assert.Equal(3, _document.Paragraphs[0].Range.ListFormat.List.ListParagraphs.Count);
    }

    [Fact]
    public void AnOutlineHasLettersAndRomanNumbersOnItsLevels()
    {
        AddNumbering(OutlineTemplate(0));
        Number(0, 1, 0);
        Number(1, 1, 1);
        Number(2, 1, 2);
        Number(3, 1, 0);

        string[] strings = Enumerable.Range(0, 4).Select(i => _document.Paragraphs[i].Range.ListFormat.ListString).ToArray();
        int[] levels = Enumerable.Range(0, 4).Select(i => _document.Paragraphs[i].Range.ListFormat.ListLevelNumber).ToArray();

        Assert.Equal(new[] { "1)", "a)", "i)", "2)" }, strings);
        Assert.Equal(new[] { 1, 2, 3, 1 }, levels);
        Assert.Equal(ListType.OutlineNumbering, _document.Paragraphs[0].Range.ListFormat.ListType);
    }

    [Fact]
    public void ADeeperLevelStartsAgainAndTextHasThePatternOfItsLevel()
    {
        AddNumbering(OutlineTemplate(0));
        Number(0, 1, 0);
        Number(1, 1, 4);
        Number(2, 1, 0);
        Number(3, 1, 4);

        Assert.Equal("(a)", _document.Paragraphs[1].Range.ListFormat.ListString);
        Assert.Equal(1, _document.Paragraphs[1].Range.ListFormat.ListValue);
        Assert.Equal("(a)", _document.Paragraphs[3].Range.ListFormat.ListString);
    }

    [Fact]
    public void ListsAreToldApartByTheirNumbering()
    {
        AddNumbering(NumberTemplate(0), NumberTemplate(1));
        Number(0, 1);
        Number(1, 1);
        Number(3, 2);

        Assert.Equal(2, _document.Lists.Count);
        Assert.Equal("1.", _document.Paragraphs[3].Range.ListFormat.ListString);
        Assert.Equal(2, _document.Lists[0].CountNumberedItems());
        Assert.NotEqual(_document.Lists[0], _document.Lists[1]);
        Assert.True(_document.Paragraphs[0].Range.ListFormat.SingleList);
        Assert.False(_document.Content.ListFormat.SingleList);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Lists[2]);
    }

    [Fact]
    public void ALevelReadsBackItsFormatting()
    {
        AddNumbering(NumberTemplate(0));

        ListTemplateWrapper template = _document.ListTemplates[0];
        ListLevelWrapper level = template.ListLevels[0];

        Assert.False(template.OutlineNumbered);
        Assert.Equal(string.Empty, template.Name);
        Assert.Equal(1, template.ListLevels.Count);
        Assert.Equal("%1.", level.NumberFormat);
        Assert.Equal(ListNumberStyle.Arabic, level.NumberStyle);
        Assert.Equal(1, level.StartAt);
        Assert.Equal(18, level.NumberPosition);
        Assert.Equal(36, level.TextPosition);
        Assert.Equal(9999999, level.TabPosition);
        Assert.Equal(TrailingCharacter.Tab, level.TrailingCharacter);
        Assert.Equal(ListLevelAlignment.Left, level.Alignment);
        Assert.Equal(0, level.ResetOnHigher);
        Assert.Equal(string.Empty, level.LinkedStyle);
        Assert.Throws<ArgumentOutOfRangeException>(() => template.ListLevels[1]);
    }

    [Fact]
    public void ALevelCanBeChanged()
    {
        AddNumbering(OutlineTemplate(0));
        Number(0, 1, 0);
        Number(1, 1, 1);
        ListLevelWrapper first = _document.ListTemplates[0].ListLevels[0];

        first.NumberStyle = ListNumberStyle.UppercaseRoman;
        first.NumberFormat = "%1.";
        first.StartAt = 3;
        first.TextPosition = 54;
        first.NumberPosition = 18;
        first.TrailingCharacter = TrailingCharacter.Space;
        first.Alignment = ListLevelAlignment.Right;

        Assert.Equal("III.", _document.Paragraphs[0].Range.ListFormat.ListString);
        Assert.Equal(54, first.TextPosition);
        Assert.Equal(18, first.NumberPosition);
        Assert.Equal(TrailingCharacter.Space, first.TrailingCharacter);
        Assert.Equal(ListLevelAlignment.Right, first.Alignment);
        Assert.Equal(1, _document.ListTemplates[0].ListLevels[1].ResetOnHigher);
        Assert.True(_document.ListTemplates[0].OutlineNumbered);
    }
}
