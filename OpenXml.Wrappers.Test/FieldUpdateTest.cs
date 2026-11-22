using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class FieldUpdateTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public FieldUpdateTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "Start ";
        _document.BuiltInDocumentProperties.Title = "My title";
        _document.BuiltInDocumentProperties.Author = "Pat Writer";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private FieldWrapper Add(string code)
    {
        return _document.Fields.Add(_document.Paragraphs.Last.Range, FieldType.Empty, code, false);
    }

    //The flag that tells Word to update the field when the document is opened.
    private static bool IsDirty(FieldWrapper field)
    {
        //A field that was worked out when it was added is one element, without field characters.
        FieldChar begin = field.Range.StartElement.Descendants<FieldChar>().FirstOrDefault();
        return begin != null && begin.Dirty != null && begin.Dirty.HasValue && begin.Dirty.Value;
    }

    private void AssertValid()
    {
        _document.Save();
        IList<string> errors = DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    [Fact]
    public void ANewFieldShowsItsResultAtOnceWhenItCanBeWorkedOut()
    {
        FieldWrapper title = Add("TITLE");
        FieldWrapper date = Add("DATE \\@ \"yyyy-MM-dd\"");

        Assert.Equal("My title", title.Result.Text);
        Assert.Equal(DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), date.Result.Text);
        Assert.False(IsDirty(title));
        AssertValid();
    }

    [Fact]
    public void FieldsThatDependOnLayoutAreLeftForWordToUpdate()
    {
        FieldWrapper page = Add("PAGE");
        FieldWrapper pages = Add("NUMPAGES");

        Assert.False(page.Update());
        Assert.False(pages.Update());
        Assert.Equal("", page.Result.Text);
        Assert.True(IsDirty(page));
        Assert.Equal(2, _document.Fields.Update());
    }

    [Fact]
    public void ADatePictureIsFormattedTheWayWordDoes()
    {
        FieldWrapper field = Add("DATE");
        DateTime now = DateTime.Now;

        field.Code.Text = "DATE \\@ \"dddd, d MMMM yyyy 'at' HH:mm\"";
        Assert.True(field.Update());

        string text = field.Result.Text;
        Assert.Equal(
            CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(now.DayOfWeek) + ", " + now.Day.ToString(CultureInfo.InvariantCulture) + " "
            + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(now.Month) + " " + now.Year.ToString(CultureInfo.InvariantCulture) + " at ",
            text.Substring(0, text.Length - 5));
        Assert.Matches("^[0-2][0-9]:[0-5][0-9]$", text.Substring(text.Length - 5));
    }

    [Theory]
    [InlineData("h:mm AM/PM", "^[0-9]{1,2}:[0-5][0-9] (AM|PM)$")]
    [InlineData("h:mm am/pm", "^[0-9]{1,2}:[0-5][0-9] (am|pm)$")]
    [InlineData("yy/M/d", "^[0-9]{2}/[0-9]{1,2}/[0-9]{1,2}$")]
    [InlineData("MMM", "^[^0-9]+$")]
    public void MoreDatePictures(string picture, string pattern)
    {
        FieldWrapper field = Add("TIME \\@ \"" + picture + "\"");

        Assert.Matches(pattern, field.Result.Text);
    }

    [Fact]
    public void CaseSwitchesChangeTheResult()
    {
        Assert.Equal("MY TITLE", Add("TITLE \\* Upper").Result.Text);
        Assert.Equal("my title", Add("TITLE \\* Lower").Result.Text);
        Assert.Equal("My title", Add("TITLE \\* FirstCap").Result.Text);
        Assert.Equal("My Title", Add("TITLE \\* Caps").Result.Text);
    }

    [Fact]
    public void DocumentPropertiesCanBeFound()
    {
        _document.CustomDocumentProperties.Add("Project", DocumentPropertyType.Text, "Falcon");
        _document.CustomDocumentProperties.Add("Count", DocumentPropertyType.Number, 12);

        Assert.Equal("Pat Writer", Add("AUTHOR").Result.Text);
        Assert.Equal("Pat Writer", Add("DOCPROPERTY Author").Result.Text);
        Assert.Equal("Falcon", Add("DOCPROPERTY Project").Result.Text);
        Assert.Equal("12", Add("DOCPROPERTY Count").Result.Text);
        FieldWrapper unknown = Add("DOCPROPERTY Nothing");
        Assert.False(unknown.Update());
        Assert.StartsWith("Error!", unknown.Result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void TheFileNameIsKnownFromWhereTheDocumentWasOpenedOrSaved()
    {
        Assert.Equal(Path.GetFileName(_path), Add("FILENAME").Result.Text);
        Assert.Equal(Path.GetFullPath(_path), Add("FILENAME \\p").Result.Text);

        string copyPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (WordprocessingDocumentWrapper copy = _document.SaveAs(copyPath))
            {
                FieldWrapper field = copy.Fields.First(f => f.Code.Text == "FILENAME");
                field.Update();
                Assert.Equal(Path.GetFileName(copyPath), field.Result.Text);
            }
        }
        finally
        {
            File.Delete(copyPath);
        }
    }

    [Fact]
    public void WordsAndCharactersAreCounted()
    {
        _document.Content.Text = "one two  three\rfour";
        FieldWrapper words = _document.Fields.Add(_document.Paragraphs.Last.Range, FieldType.NumWords, null, false);
        FieldWrapper characters = _document.Fields.Add(_document.Paragraphs.Last.Range, FieldType.NumChars, null, false);

        Assert.Equal("4", words.Result.Text);
        //The result of the words field (one character) is in the document by now.
        Assert.Equal("16", characters.Result.Text);
    }

    [Fact]
    public void SequenceNumbersCountThroughTheDocument()
    {
        FieldWrapper first = Add("SEQ Figure");
        FieldWrapper other = Add("SEQ Table");
        FieldWrapper second = Add("SEQ Figure");
        FieldWrapper repeat = Add("SEQ Figure \\c");
        FieldWrapper reset = Add("SEQ Figure \\r 10");
        FieldWrapper next = Add("SEQ Figure");
        FieldWrapper roman = Add("SEQ Figure \\* roman");
        FieldWrapper letters = Add("SEQ Figure \\* ALPHABETIC");
        FieldWrapper hidden = Add("SEQ Figure \\h");

        Assert.Equal("1", first.Result.Text);
        Assert.Equal("1", other.Result.Text);
        Assert.Equal("2", second.Result.Text);
        Assert.Equal("2", repeat.Result.Text);
        Assert.Equal("10", reset.Result.Text);
        Assert.Equal("11", next.Result.Text);
        Assert.Equal("xii", roman.Result.Text);
        Assert.Equal("M", letters.Result.Text);
        Assert.Equal("", hidden.Result.Text);

        _document.Fields.Update();
        Assert.Equal("2", second.Result.Text);
        Assert.Equal("11", next.Result.Text);
    }

    //Documents of headings and SEQ fields with \s, and the results Word gave for them (#93). Each line is "style|text|code|result".
    [Theory]
    [InlineData(
        "Heading 1|A||", "Normal|f1 |SEQ F \\s 1|1", "Normal|f2 |SEQ F|2",
        "Heading 1|B||", "Normal|f3 |SEQ F|3", "Normal|f4 |SEQ F \\s 1|4",
        "Heading 2|C||", "Normal|f5 |SEQ F \\s 2|1", "Normal|f6 |SEQ F \\s 1|2",
        "Heading 3|D||", "Normal|f7 |SEQ F \\s 2|3",
        "Heading 1|E |SEQ F \\s 1|1", "Normal|f9 |SEQ F \\s 1|2",
        "Heading 1|||", "Normal|f10 |SEQ F \\s 1|1",
        "Normal|f11 |SEQ F \\s 1 \\r 5|2", "Normal|f12 |SEQ F \\s 1 \\c|3",
        "Heading 1|G||", "Normal|f13 |SEQ F \\s 1 \\c|1", "Normal|f14 |SEQ G \\s 1|1")]
    [InlineData(
        "Heading 1|A||", "Normal|a1 |SEQ F|1", "Normal|a2 |SEQ F \\r 5|5", "Normal|a3 |SEQ F \\c|5",
        "Normal|b1 |SEQ F \\r 7 \\s 1|7", "Normal|b2 |SEQ F \\c \\s 1|7", "Normal|b3 |SEQ F \\s 1 \\c|8", "Normal|b4 |SEQ F \\s 1 \\r 9|9",
        "Normal|b5 |SEQ F \\s 1 \\h|", "Normal|b6 |SEQ F \\s 1 \\* ROMAN|XI", "Normal|b7 |SEQ F \\s 1 \\n|12",
        "Heading 1|B||", "Normal|c1 |SEQ F \\s 1 \\r 3|1", "Normal|c2 |SEQ F \\s 1|2")]
    [InlineData(
        "Heading 1|A||", "Normal|x1 |SEQ F|1", "Normal|x2 |SEQ F|2",
        "Heading 1|B||", "Normal|y1 |SEQ F \\n \\s 1|1", "Normal|y2 |SEQ F \\c \\r 9|1", "Normal|y3 |SEQ F \\r 20 \\c|20",
        "Heading 1|C||", "Normal|z1 |SEQ F \\c \\s 1|0",
        "Heading 1|D||", "Normal|w1 |SEQ F \\r 4 \\s 1|4",
        "Heading 1|E||", "Normal|v1 |SEQ F \\h \\s 1|", "Normal|v2 |SEQ F \\s 1 \\n|2")]
    public void SequenceNumbersStartAgainAfterAHeadingWithTheRestartSwitch(params string[] lines)
    {
        System.Collections.Generic.List<(FieldWrapper Field, string Expected)> fields = new System.Collections.Generic.List<(FieldWrapper, string)>();
        foreach (string line in lines)
        {
            string[] parts = line.Split('|');
            ParagraphWrapper paragraph = _document.Paragraphs.Add();
            paragraph.Range.Text = parts[1];
            paragraph.Style = parts[0];
            if (parts[2].Length > 0)
            {
                fields.Add((Add(parts[2]), parts[3]));
            }
        }

        _document.Fields.Update();

        foreach ((FieldWrapper field, string expected) in fields)
        {
            Assert.True(expected == field.Result.Text, string.Format(CultureInfo.InvariantCulture, "{0}: expected '{1}', got '{2}'", field.Code.Text, expected, field.Result.Text));
        }
        AssertValid();
    }

    [Fact]
    public void ARefFieldShowsTheTextOfItsBookmark()
    {
        _document.Content.Text = "The quick brown fox";
        _document.Bookmarks.Add("animal", _document.Range(16, 19));

        FieldWrapper reference = Add("REF animal");
        FieldWrapper missing = Add("REF nothing");

        Assert.Equal("fox", reference.Result.Text);
        Assert.False(missing.Update());
        Assert.Equal("Error! Reference source not found.", missing.Result.Text);
    }

    [Theory]
    [InlineData("IF 1 = 1 \"yes\" \"no\"", "yes")]
    [InlineData("IF 2 > 3 \"yes\" \"no\"", "no")]
    [InlineData("IF 10 >= 9 \"yes\" \"no\"", "yes")]
    [InlineData("IF \"Cat\" = \"cat\" \"yes\" \"no\"", "yes")]
    [InlineData("IF \"Cattle\" = \"Cat*\" \"yes\" \"no\"", "yes")]
    [InlineData("IF a <> b \"yes\" \"no\"", "yes")]
    [InlineData("IF a < b \"yes\" \"no\"", "yes")]
    [InlineData("IF 1 = 2 \"yes\"", "")]
    public void AnIfFieldComparesItsOperands(string code, string expected)
    {
        Assert.Equal(expected, Add(code).Result.Text);
    }

    [Fact]
    public void AQuoteFieldShowsItsText()
    {
        Assert.Equal("Hello there", Add("QUOTE \"Hello there\"").Result.Text);
    }

    [Fact]
    public void ALockedFieldIsNotUpdated()
    {
        FieldWrapper field = Add("TITLE");
        field.Locked = true;
        _document.BuiltInDocumentProperties.Title = "New title";

        Assert.True(field.Locked);
        Assert.False(field.Update());
        Assert.Equal("My title", field.Result.Text);

        field.Locked = false;
        Assert.True(field.Update());
        Assert.Equal("New title", field.Result.Text);
        AssertValid();
    }

    [Fact]
    public void FieldsCanBeLockedAndUnlockedTogether()
    {
        Add("TITLE");
        Add("AUTHOR");
        Assert.False(_document.Fields.Locked);

        _document.Fields.Lock();
        Assert.True(_document.Fields.Locked);
        _document.Fields[0].Locked = false;
        Assert.Null(_document.Fields.Locked);
        _document.Fields.Unlock();
        Assert.False(_document.Fields.Locked);
    }

    [Fact]
    public void AFieldKnowsItsPlaceAndNeighbours()
    {
        FieldWrapper first = Add("TITLE");
        FieldWrapper second = Add("AUTHOR");
        FieldWrapper third = Add("SUBJECT");

        Assert.Equal(0, first.Index);
        Assert.Equal(2, third.Index);
        Assert.Equal(second, first.Next());
        Assert.Equal(first, second.Previous());
        Assert.Null(first.Previous());
        Assert.Null(third.Next());
    }

    [Fact]
    public void UnlinkLeavesTheResultAsText()
    {
        FieldWrapper title = Add("TITLE");
        Add("AUTHOR");

        title.Unlink();

        Assert.Equal(1, _document.Fields.Count);
        Assert.Contains("My title", _document.Content.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("TITLE", _document.WordprocessingDocument.MainDocumentPart.Document.InnerXml, StringComparison.Ordinal);
        AssertValid();
    }

    [Fact]
    public void UnlinkingASimpleFieldKeepsItsRuns()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a")), new SimpleField(new Run(new Text("RESULT"))) { Instruction = " TITLE " }, new Run(new Text("b")));
        FieldWrapper field = new RangeWrapper(paragraph).Fields.Single();

        field.Unlink();

        Assert.Equal("aRESULTb", paragraph.InnerText);
        Assert.Empty(paragraph.Elements<SimpleField>());
    }

    [Fact]
    public void FieldTypesHaveTheValuesOfWord()
    {
        Assert.Equal(15, (int)FieldType.Title);
        Assert.Equal(85, (int)FieldType.DocProperty);
        Assert.Equal(FieldType.DocProperty, Add("DOCPROPERTY Title").Type);
        Assert.Equal(FieldType.Formula, new RangeWrapper(new Paragraph(new SimpleField() { Instruction = " = 1+1 " })).Fields.Single().Type);
        Assert.Equal(FieldType.Quote, Add("QUOTE x").Type);
    }

    [Fact]
    public void PreserveFormattingAddsMergeFormat()
    {
        FieldWrapper preserved = _document.Fields.Add(_document.Paragraphs.Last.Range, FieldType.Title);
        FieldWrapper plain = _document.Fields.Add(_document.Paragraphs.Last.Range, FieldType.Subject, null, false);

        Assert.Equal("TITLE  \\* MERGEFORMAT", preserved.Code.Text);
        Assert.Equal("SUBJECT", plain.Code.Text);
    }
}
