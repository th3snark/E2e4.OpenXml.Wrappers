using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// REF fields with the list-number switches (\n, \r, \w, \t, \p) give what Word gives (#95). The documents are those of
/// tools/WordReference/Measure-RefNumbers.ps1, and TestFiles/RefNumbers.txt holds what Word showed for each field in them.
/// </summary>
public class FieldReferenceNumberTest
{
    private const string Results = @"TestFiles\RefNumbers.txt";

    //The levels of each template (number format and style), as in Measure-RefNumbers.ps1.
    private static readonly Dictionary<int, (string Format, ListNumberStyle Style)[]> Templates = new Dictionary<int, (string, ListNumberStyle)[]>
    {
        [1] = new[] { ("Chapter %1", ListNumberStyle.Arabic), ("(%2)", ListNumberStyle.LowercaseLetter), ("%3.", ListNumberStyle.LowercaseRoman) },
        [2] = new[] { ("%1.", ListNumberStyle.Arabic), ("%1.%2.", ListNumberStyle.Arabic), ("%1.%2.%3.", ListNumberStyle.Arabic) },
        [3] = new[] { ("%1)", ListNumberStyle.Arabic), ("%2)", ListNumberStyle.LowercaseLetter), ("%3)", ListNumberStyle.LowercaseRoman) },
        [4] = new[] { ("%1.", ListNumberStyle.Arabic), ("%2.", ListNumberStyle.LowercaseLetter), ("%3.", ListNumberStyle.LowercaseRoman) },
        [5] = new[] { ("Part %1:", ListNumberStyle.Arabic), ("Section %1-%2", ListNumberStyle.Arabic), ("Clause %3 of", ListNumberStyle.Arabic) },
    };

    private static readonly Dictionary<string, (string Style, string Text)[]> Layouts = new Dictionary<string, (string, string)[]>
    {
        ["A"] = new[]
        {
            ("Normal", "Start"), ("Heading 1", "One"), ("Heading 2", "OneA"), ("Heading 3", "OneAi"), ("Heading 3", "OneAii"),
            ("Heading 2", "OneB"), ("Heading 1", "Two"), ("Heading 2", "TwoA"), ("Heading 3", "TwoAi"),
        },
        ["B"] = new[]
        {
            ("Normal", "Start"), ("Heading 1", "One"), ("Normal", "BodyOne"), ("Heading 2", "OneA"), ("Heading 3", "OneAi"),
            ("Normal", "BodyOneAi"), ("Normal", "Plain"), ("Heading 2", "OneB"), ("Heading 1", "Two"), ("Normal", "BodyTwo"),
        },
    };

    private static List<string[]> ReadResults()
    {
        return File.ReadAllLines(Results).Where(l => l.Length > 0 && l[0] != '#').Select(l => l.Split('\t')).ToList();
    }

    public static IEnumerable<object[]> Documents()
    {
        return ReadResults().Select(r => (r[0], r[1])).Distinct().Select(d => new object[] { int.Parse(d.Item1, CultureInfo.InvariantCulture), d.Item2 });
    }

    [Theory]
    [MemberData(nameof(Documents))]
    public void ListNumberSwitchesGiveWhatWordGives(int template, string layout)
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        (string Style, string Text)[] paragraphs = Layouts[layout];
        document.Content.Text = string.Join("\r", paragraphs.Select(p => p.Text));

        ListTemplateWrapper list = document.ListTemplates.Add(true);
        for (int i = 0; i < 3; i++)
        {
            list.ListLevels[i].NumberFormat = Templates[template][i].Format;
            list.ListLevels[i].NumberStyle = Templates[template][i].Style;
            list.ListLevels[i].LinkedStyle = "Heading " + (i + 1).ToString(CultureInfo.InvariantCulture);
        }
        document.Paragraphs[1].Range.ListFormat.ApplyListTemplateWithLevel(list);
        for (int k = 0; k < paragraphs.Length; k++)
        {
            document.Paragraphs[k].Style = paragraphs[k].Style;
        }
        for (int k = 0; k < paragraphs.Length; k++)
        {
            RangeWrapper text = document.Paragraphs[k].Range;
            document.Bookmarks.Add("B" + (k + 1).ToString(CultureInfo.InvariantCulture), document.Range(text.Start, text.End - 1));
        }

        List<string> differences = new List<string>();
        int count = 0;
        foreach (string[] row in ReadResults().Where(r => r[0] == template.ToString(CultureInfo.InvariantCulture) && r[1] == layout))
        {
            int host = Array.FindIndex(paragraphs, p => p.Text == row[2]);
            int target = Array.FindIndex(paragraphs, p => p.Text == row[3]) + 1;
            RangeWrapper at = document.Paragraphs[host].Range;
            FieldWrapper field = document.Fields.Add(document.Range(at.End - 1, at.End - 1), FieldType.Empty,
                "REF B" + target.ToString(CultureInfo.InvariantCulture) + " " + row[4], false);
            field.Update();
            if (field.Result.Text != row[5])
            {
                differences.Add(string.Format(CultureInfo.InvariantCulture, "from {0} to {1} {2}: Word '{3}', library '{4}'", row[2], row[3], row[4], row[5], field.Result.Text));
            }
            field.Delete();
            count++;
        }

        Assert.True(count > 0);
        Assert.True(differences.Count == 0, string.Join(Environment.NewLine, differences));
    }
}
