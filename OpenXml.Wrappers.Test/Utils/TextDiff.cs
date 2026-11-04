using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OpenXml.Wrappers.Test.Utils;

/// <summary>
/// A line diff for failure messages: lines only in the expected text start with <c>-</c>, lines only in the actual text with <c>+</c>, and a
/// few unchanged lines around each change are kept for context.
/// </summary>
public static class TextDiff
{
    public static string Lines(string expected, string actual, int context = 3, int maxLines = 120)
    {
        string[] a = Split(expected);
        string[] b = Split(actual);

        //The longest common subsequence, by dynamic programming from the end.
        int[,] lengths = new int[a.Length + 1, b.Length + 1];
        for (int i = a.Length - 1; i >= 0; i--)
        {
            for (int j = b.Length - 1; j >= 0; j--)
            {
                lengths[i, j] = a[i] == b[j] ? lengths[i + 1, j + 1] + 1 : Math.Max(lengths[i + 1, j], lengths[i, j + 1]);
            }
        }

        List<KeyValuePair<char, string>> edits = new List<KeyValuePair<char, string>>();
        int x = 0;
        int y = 0;
        while (x < a.Length || y < b.Length)
        {
            if (x < a.Length && y < b.Length && a[x] == b[y])
            {
                edits.Add(new KeyValuePair<char, string>(' ', a[x++]));
                y++;
            }
            else if (y < b.Length && (x == a.Length || lengths[x, y + 1] >= lengths[x + 1, y]))
            {
                edits.Add(new KeyValuePair<char, string>('+', b[y++]));
            }
            else
            {
                edits.Add(new KeyValuePair<char, string>('-', a[x++]));
            }
        }

        bool[] shown = new bool[edits.Count];
        for (int i = 0; i < edits.Count; i++)
        {
            if (edits[i].Key != ' ')
            {
                for (int k = Math.Max(0, i - context); k <= Math.Min(edits.Count - 1, i + context); k++)
                {
                    shown[k] = true;
                }
            }
        }

        StringBuilder builder = new StringBuilder();
        int written = 0;
        for (int i = 0; i < edits.Count; i++)
        {
            if (!shown[i])
            {
                continue;
            }
            if (i > 0 && !shown[i - 1])
            {
                builder.Append("  ...\n");
            }
            if (written == maxLines)
            {
                builder.Append(string.Format(CultureInfo.InvariantCulture, "  ... ({0} more changed lines)\n", edits.Skip(i).Count(e => e.Key != ' ')));
                break;
            }
            builder.Append(edits[i].Key).Append(' ').Append(edits[i].Value).Append('\n');
            written++;
        }
        return builder.ToString();
    }

    private static string[] Split(string text)
    {
        return text.Replace("\r\n", "\n").TrimEnd('\n').Split('\n');
    }
}
