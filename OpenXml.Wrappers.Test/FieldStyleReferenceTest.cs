using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// STYLEREF fields in the main text give what Word gives (#94). The document is that of tools/WordReference/Measure-StyleRefs.ps1, and
/// TestFiles/StyleRefs.txt holds what Word showed for each field in it.
/// </summary>
public class FieldStyleReferenceTest
{
    private const string Results = @"TestFiles\StyleRefs.txt";

    private static readonly (string Style, string Text)[] Paragraphs =
    {
        ("Normal", "Start"), ("Heading 1", "Intro"), ("Heading 2", "Alpha"), ("Normal", "BodyAlpha"), ("Heading 1", "Second"),
        ("Normal", "BodySecond"), ("Heading 2", "Beta"), ("Heading 2", string.Empty), ("Normal", "BodyEmpty"), ("Normal", "Strong words here"),
    };

    //Codes the library leaves for Word: a character style, and \s, whose rule is not known.
    private static readonly string[] LeftForWord = { "STYLEREF Strong", "STYLEREF \"Heading 1\" \\s" };

    [Fact]
    public void StyleReferencesInTheMainTextGiveWhatWordGives()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Content.Text = string.Join("\r", Paragraphs.Select(p => p.Text));

        ListTemplateWrapper list = document.ListTemplates.Add(true);
        (string Format, ListNumberStyle Style)[] levels = { ("Chapter %1", ListNumberStyle.Arabic), ("(%2)", ListNumberStyle.LowercaseLetter), ("%3.", ListNumberStyle.LowercaseRoman) };
        for (int i = 0; i < levels.Length; i++)
        {
            list.ListLevels[i].NumberFormat = levels[i].Format;
            list.ListLevels[i].NumberStyle = levels[i].Style;
            list.ListLevels[i].LinkedStyle = "Heading " + (i + 1).ToString(CultureInfo.InvariantCulture);
        }
        document.Paragraphs[1].Range.ListFormat.ApplyListTemplateWithLevel(list);
        for (int k = 0; k < Paragraphs.Length; k++)
        {
            document.Paragraphs[k].Style = Paragraphs[k].Style;
        }
        RangeWrapper last = document.Paragraphs[Paragraphs.Length - 1].Range;
        document.Range(last.Start, last.Start + 6).Style = "Strong";

        List<string> differences = new List<string>();
        int count = 0;
        foreach (string[] row in File.ReadAllLines(Results).Where(l => l.Length > 0 && l[0] != '#').Select(l => l.Split('\t')))
        {
            RangeWrapper at = document.Paragraphs[Array.FindIndex(Paragraphs, p => p.Text == row[0])].Range;
            FieldWrapper field = document.Fields.Add(document.Range(at.End - 1, at.End - 1), FieldType.Empty, row[1], false);
            field.Result.Text = "?";
            bool updated = field.Update();
            string expected = LeftForWord.Contains(row[1]) ? "?" : row[2];
            if (field.Result.Text != expected || (expected == "?") == updated && !row[2].StartsWith("Error!", StringComparison.Ordinal))
            {
                differences.Add(string.Format(CultureInfo.InvariantCulture, "from {0} {1}: expected '{2}', library '{3}' (updated {4})", row[0], row[1], expected, field.Result.Text, updated));
            }
            field.Delete();
            count++;
        }

        Assert.True(count > 0);
        Assert.True(differences.Count == 0, string.Join(Environment.NewLine, differences));
    }
}
