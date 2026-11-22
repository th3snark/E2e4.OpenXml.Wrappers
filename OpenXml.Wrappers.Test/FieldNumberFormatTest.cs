using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The number formats of the <c>\*</c> switch give what Word gives (#99). TestFiles/NumberFormats.txt holds what Word showed for each code,
/// written by tools/WordReference/Measure-NumberFormats.ps1, which adds each field to an empty document on its own.
/// </summary>
public class FieldNumberFormatTest
{
    private const string Results = @"TestFiles\NumberFormats.txt";

    [Fact]
    public void NumberFormatsGiveWhatWordGives()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);

        List<string> differences = new List<string>();
        int count = 0;
        foreach (string[] row in File.ReadAllLines(Results).Where(l => l.Length > 0 && l[0] != '#').Select(l => l.Split('\t')))
        {
            FieldWrapper field = document.Fields.Add(document.Range(0, 0), FieldType.Empty, row[0], false);
            field.Result.Text = "?";
            bool updated = field.Update();
            if (field.Result.Text != row[1] || !updated)
            {
                differences.Add(string.Format(CultureInfo.InvariantCulture, "{0}: expected '{1}', library '{2}' (updated {3})", row[0], row[1], field.Result.Text, updated));
            }
            field.Delete();
            count++;
        }

        Assert.True(count > 3000);
        Assert.True(differences.Count == 0, differences.Count.ToString(CultureInfo.InvariantCulture) + " differences:" + Environment.NewLine + string.Join(Environment.NewLine, differences.Take(60)));
    }

    //The case formats, and what Word 16 gave for them: FirstCap only makes the first character after any spaces a capital (if it is a letter),
    //and Caps keeps a word in capitals when the text has lowercase letters too. The names are not case sensitive.
    [Theory]
    [InlineData(@"QUOTE ""hELLO wORLD"" \* FirstCap", "HELLO wORLD")]
    [InlineData(@"QUOTE ""hELLO wORLD"" \* Caps", "Hello World")]
    [InlineData(@"QUOTE ""hELLO wORLD"" \* Upper", "HELLO WORLD")]
    [InlineData(@"QUOTE ""hELLO wORLD"" \* Lower", "hello world")]
    [InlineData(@"QUOTE ""hELLO wORLD"" \* firstcap", "HELLO wORLD")]
    [InlineData(@"QUOTE ""hELLO wORLD"" \* CAPS", "Hello World")]
    [InlineData(@"QUOTE ""hELLO wORLD"" \* upper", "HELLO WORLD")]
    [InlineData(@"QUOTE ""HELLO WORLD"" \* FirstCap", "HELLO WORLD")]
    [InlineData(@"QUOTE ""HELLO WORLD"" \* Caps", "Hello World")]
    [InlineData(@"QUOTE ""HELLO WORLD"" \* Lower", "hello world")]
    [InlineData(@"QUOTE ""hello world"" \* FirstCap", "Hello world")]
    [InlineData(@"QUOTE ""hello world"" \* Caps", "Hello World")]
    [InlineData(@"QUOTE ""the QUICK Fox"" \* FirstCap", "The QUICK Fox")]
    [InlineData(@"QUOTE ""the QUICK Fox"" \* Caps", "The QUICK Fox")]
    [InlineData(@"QUOTE ""the QUICK Fox"" \* Upper", "THE QUICK FOX")]
    [InlineData(@"QUOTE ""the QUICK Fox"" \* Lower", "the quick fox")]
    [InlineData(@"QUOTE "" leading space"" \* FirstCap", " Leading space")]
    [InlineData(@"QUOTE "" leading space"" \* Caps", " Leading Space")]
    [InlineData(@"QUOTE ""1st PLACE"" \* FirstCap", "1st PLACE")]
    [InlineData(@"QUOTE ""1st PLACE"" \* Caps", "1st PLACE")]
    [InlineData(@"QUOTE ""1st PLACE"" \* Upper", "1ST PLACE")]
    [InlineData(@"QUOTE ""éCOLE Normale"" \* FirstCap", "ÉCOLE Normale")]
    [InlineData(@"QUOTE ""éCOLE Normale"" \* Caps", "École Normale")]
    [InlineData(@"QUOTE ""éCOLE Normale"" \* Upper", "ÉCOLE NORMALE")]
    [InlineData(@"QUOTE ""éCOLE Normale"" \* Lower", "école normale")]
    [InlineData(@"QUOTE ""x"" \* FirstCap", "X")]
    [InlineData(@"QUOTE ""x"" \* Caps", "X")]
    public void CaseFormatsGiveWhatWordGives(string code, string expected)
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        FieldWrapper field = document.Fields.Add(document.Range(0, 0), FieldType.Empty, code, false);

        Assert.True(field.Update());
        Assert.Equal(expected, field.Result.Text);
    }

    //A number format before the picture switch is an error in Word ("Picture switch must be first formatting switch."), which is left for
    //Word, and so is a number with an exponent, which Word reads in its own way.
    [Theory]
    [InlineData("= 21 \\* CardText \\# \"0.00\"")]
    [InlineData("QUOTE \"3e2\" \\* CardText")]
    public void FormatsNotWorkedOutAreLeftForWord(string code)
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        FieldWrapper field = document.Fields.Add(document.Range(0, 0), FieldType.Empty, code, false);
        field.Result.Text = "?";

        Assert.False(field.Update());
        Assert.Equal("?", field.Result.Text);
    }
}
