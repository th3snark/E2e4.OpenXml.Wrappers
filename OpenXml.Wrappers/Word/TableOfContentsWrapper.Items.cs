using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class TableOfContentsWrapperCollection
{
    /// <summary>One entry of a table of contents or figures: a heading, a caption or a TC field, with the text and level of the entry.</summary>
    internal sealed class TocItem
    {
        public Paragraph Paragraph;
        public int Level;
        public string Text;

        //For a caption listed without its label and number: how many characters of the paragraph come before the entry.
        public int EntryOffset;

        //For a TC field: its first and last runs, the elements between, the entry and what follows it in the code.
        public Run TcBegin;
        public Run TcEnd;
        public string TcEntry;
        public string TcRest;
    }

    //The entries of a table in the order of the text: headings (or captions) and, for a table that uses fields, TC fields.
    internal static List<TocItem> CollectItems(OpenXmlElement root, TocSettings settings, Styles styles)
    {
        List<TocItem> items = new List<TocItem>();
        string tableId = string.IsNullOrEmpty(settings.TableId) ? "C" : settings.TableId;
        foreach (Paragraph paragraph in root.Descendants<Paragraph>().ToList())
        {
            if (paragraph.Ancestors<Hyperlink>().Any())
            {
                continue;
            }
            if (settings.Caption != null)
            {
                if (IsCaption(paragraph, settings.Caption))
                {
                    string captionText = CaptionText(paragraph, settings);
                    items.Add(new TocItem { Paragraph = paragraph, Level = 1, Text = captionText, EntryOffset = TextOf(paragraph).Length - captionText.Length });
                }
                continue;
            }

            int level = LevelOf(paragraph, settings, styles);
            if (level > 0)
            {
                items.Add(new TocItem { Paragraph = paragraph, Level = level, Text = TextOf(paragraph) });
            }
            if (settings.UseFields)
            {
                foreach (TocItem tc in TcFields(paragraph).Where(t => string.Equals(TcId(t), tableId, StringComparison.OrdinalIgnoreCase)))
                {
                    items.Add(tc);
                }
            }
        }
        return items;
    }

    //Makes a table of figures: a table of contents whose entries are the captions with a label.
    internal TableOfContentsWrapper AddCaptionTable(RangeWrapper range, string caption, bool includeLabel, bool useHeadingStyles, int upperHeadingLevel,
        int lowerHeadingLevel, bool useFields, string tableID, bool rightAlignPageNumbers, bool includePageNumbers, string addedStyles, bool useHyperlinks,
        bool hidePageNumbersInWeb)
    {
        if (upperHeadingLevel < 1 || upperHeadingLevel > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(upperHeadingLevel));
        }
        if (lowerHeadingLevel < upperHeadingLevel || lowerHeadingLevel > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(lowerHeadingLevel));
        }

        RangeWrapper at = range.Duplicate();
        at.Collapse(CollapseDirection.Start);
        Paragraph paragraph = at.Paragraphs[0].Paragraph;
        if (at.Start != new RangeWrapper(paragraph).Start)
        {
            throw new NotSupportedException("A table of figures can be added only at the start of a paragraph.");
        }

        TocSettings settings = new TocSettings
        {
            UseHeadingStyles = useHeadingStyles,
            Upper = upperHeadingLevel,
            Lower = lowerHeadingLevel,
            UseFields = useFields,
            TableId = string.IsNullOrEmpty(tableID) ? null : tableID,
            RightAlign = rightAlignPageNumbers,
            IncludePageNumbers = includePageNumbers,
            UseHyperlinks = useHyperlinks,
            HidePageNumbersInWeb = hidePageNumbersInWeb,
            AddedStyles = TocSettings.ParseStyles(addedStyles),
            Caption = caption,
            IncludeLabel = includeLabel
        };
        return Build(paragraph, settings, TabLeader.Dots, new Dictionary<string, string>());
    }

    /// <summary>
    /// Marks text as an entry of a table of contents that uses fields, as <c>TablesOfContents.MarkEntry</c> in Word: a <c>TC</c> field
    /// after the range. A table with the same identifier lists the entry when it is made or updated.
    /// </summary>
    /// <param name="range">The text the entry is for; the field goes after it.</param>
    /// <param name="entry">The text of the entry.</param>
    /// <param name="entryAutoText">An AutoText entry for the text; not supported.</param>
    /// <param name="tableID">The identifier of the table the entry is for; C when not given.</param>
    /// <param name="level">The level of the entry (1 to 9).</param>
    /// <exception cref="ArgumentException">Neither an entry nor an AutoText entry is given.</exception>
    /// <exception cref="NotSupportedException"><paramref name="entryAutoText"/> is given.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The level is not from 1 to 9.</exception>
    public FieldWrapper MarkEntry(RangeWrapper range, string entry = null, string entryAutoText = null, string tableID = null, int level = 1)
    {
        Guard.ThrowIfNull(range);
        if (!string.IsNullOrEmpty(entryAutoText))
        {
            throw new NotSupportedException("AutoText entries are not supported.");
        }
        if (string.IsNullOrEmpty(entry))
        {
            throw new ArgumentException("An entry must be supplied.", nameof(entry));
        }
        if (level < 1 || level > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(level));
        }

        RangeWrapper at = range.Duplicate();
        at.Collapse(CollapseDirection.End);
        FieldChar begin = new FieldChar { FieldCharType = FieldCharValues.Begin };
        string code = " TC \"" + entry + "\" \\f " + (string.IsNullOrEmpty(tableID) ? "C" : tableID) + " \\l \"" + level.ToString(CultureInfo.InvariantCulture) + "\" ";
        at.InsertNoteReference(target => new List<Run>
        {
            new Run(begin),
            new Run(new FieldCode(code) { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End })
        });
        return new FieldWrapper(begin);
    }
    //A paragraph is a caption of a label when it has a SEQ field of that label.
    private static bool IsCaption(Paragraph paragraph, string label)
    {
        foreach (SimpleField field in paragraph.Descendants<SimpleField>())
        {
            if (Regex.IsMatch(field.Instruction?.Value ?? string.Empty, "^\\s*SEQ\\s+\"?" + Regex.Escape(label) + "\"?(\\s|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return true;
            }
        }
        return paragraph.Descendants<FieldCode>().Any(c => Regex.IsMatch(c.Text ?? string.Empty, "^\\s*SEQ\\s+\"?" + Regex.Escape(label) + "\"?(\\s|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
    }

    //The text of a caption: all of it, or with the label and number left out, what follows the number without the punctuation after it.
    private static string CaptionText(Paragraph paragraph, TocSettings settings)
    {
        string text = TextOf(paragraph);
        if (settings.IncludeLabel)
        {
            return text;
        }
        SimpleField number = paragraph.Descendants<SimpleField>().FirstOrDefault(f => Regex.IsMatch(f.Instruction?.Value ?? string.Empty, "^\\s*SEQ\\s", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
        if (number == null)
        {
            return text;
        }
        string before = string.Concat(paragraph.Descendants<Text>().TakeWhile(t => !number.Descendants<Text>().Contains(t)).Select(t => t.Text));
        string numberText = string.Concat(number.Descendants<Text>().Select(t => t.Text));
        int at = before.Length + numberText.Length;
        return at <= text.Length ? text.Substring(at).TrimStart(' ', '\t', ':', '.', '-', ')', ';', ',', '–', '—') : text;
    }

    //The TC fields of a paragraph.
    internal static List<TocItem> TcFields(Paragraph paragraph)
    {
        List<TocItem> found = new List<TocItem>();
        List<OpenXmlElement> children = paragraph.ChildElements.ToList();
        for (int i = 0; i < children.Count; i++)
        {
            if (!(children[i] is Run begin) || begin.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.Begin)
            {
                continue;
            }
            int depth = 0;
            int end = -1;
            System.Text.StringBuilder code = new System.Text.StringBuilder();
            for (int j = i; j < children.Count; j++)
            {
                if (!(children[j] is Run run))
                {
                    continue;
                }
                FieldChar chars = run.GetFirstChild<FieldChar>();
                if (chars?.FieldCharType?.Value == FieldCharValues.Begin)
                {
                    depth++;
                }
                else if (chars?.FieldCharType?.Value == FieldCharValues.End)
                {
                    depth--;
                    if (depth == 0)
                    {
                        end = j;
                        break;
                    }
                }
                else if (depth == 1)
                {
                    foreach (FieldCode piece in run.Elements<FieldCode>())
                    {
                        code.Append(piece.Text);
                    }
                }
            }
            if (end < 0)
            {
                continue;
            }
            Match match = Regex.Match(code.ToString(), "^\\s*TC\\s+\"(?<entry>[^\"]*)\"(?<rest>.*)$", RegexOptions.Singleline | RegexOptions.CultureInvariant);
            if (match.Success)
            {
                found.Add(new TocItem
                {
                    Paragraph = paragraph,
                    TcBegin = begin,
                    TcEnd = (Run)children[end],
                    TcEntry = match.Groups["entry"].Value,
                    TcRest = match.Groups["rest"].Value,
                    Text = match.Groups["entry"].Value,
                    Level = TcLevel(match.Groups["rest"].Value)
                });
            }
            i = end;
        }
        return found;
    }

    private static int TcLevel(string rest)
    {
        Match level = Regex.Match(rest, "\\\\l\\s+\"?(?<l>\\d)\"?");
        return level.Success ? Math.Max(1, int.Parse(level.Groups["l"].Value, CultureInfo.InvariantCulture)) : 1;
    }

    private static string TcId(TocItem item)
    {
        Match id = Regex.Match(item.TcRest ?? string.Empty, "\\\\f\\s+(?<id>\\S+)");
        return id.Success ? id.Groups["id"].Value : "C";
    }

    private static Run CodeRun(string text)
    {
        FieldCode code = new FieldCode(text);
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
        {
            code.Space = SpaceProcessingModeValues.Preserve;
        }
        return new Run(code);
    }

    //The name of the _Toc bookmark of a paragraph, made when it has none: from its start to its end.
    internal static string EnsureHeadingBookmark(Paragraph heading, ref uint nextBookmark, ref long nextName)
    {
        BookmarkStart bookmark = heading.Elements<BookmarkStart>().FirstOrDefault(b => b.Name?.Value != null && b.Name.Value.StartsWith("_Toc", StringComparison.Ordinal));
        if (bookmark == null)
        {
            bookmark = new BookmarkStart
            {
                Id = (nextBookmark++).ToString(CultureInfo.InvariantCulture),
                Name = "_Toc" + (nextName++).ToString(CultureInfo.InvariantCulture)
            };
            BookmarkEnd bookmarkEnd = new BookmarkEnd { Id = bookmark.Id.Value };
            OpenXmlElement properties = heading.ParagraphProperties;
            if (properties != null)
            {
                properties.InsertAfterSelf(bookmark);
            }
            else
            {
                heading.PrependChild(bookmark);
            }
            heading.AppendChild(bookmarkEnd);
        }
        return bookmark.Name.Value;
    }

    //The name of the _Toc bookmark of a caption listed without its label: it starts where the entry does, which splits the run there.
    internal static string EnsureCaptionBookmark(TocItem item, ref uint nextBookmark, ref long nextName)
    {
        Paragraph paragraph = item.Paragraph;
        if (item.EntryOffset <= 0 || paragraph.Elements<BookmarkStart>().Any(b => b.Name?.Value != null && b.Name.Value.StartsWith("_Toc", StringComparison.Ordinal)))
        {
            return EnsureHeadingBookmark(paragraph, ref nextBookmark, ref nextName);
        }

        int seen = 0;
        foreach (OpenXmlElement child in paragraph.ChildElements.ToList())
        {
            int length = string.Concat(child.Descendants<Text>().Select(t => t.Text)).Length;
            if (child is Run run && run.Elements<Text>().Count() == 1 && seen <= item.EntryOffset && item.EntryOffset < seen + length)
            {
                Text text = run.GetFirstChild<Text>();
                int cut = item.EntryOffset - seen;
                Run second = (Run)run.CloneNode(true);
                text.Text = text.Text.Substring(0, cut);
                text.Space = SpaceProcessingModeValues.Preserve;
                Text secondText = second.GetFirstChild<Text>();
                secondText.Text = secondText.Text.Substring(cut);
                if (secondText.Text.Length > 0 && (char.IsWhiteSpace(secondText.Text[0]) || char.IsWhiteSpace(secondText.Text[secondText.Text.Length - 1])))
                {
                    secondText.Space = SpaceProcessingModeValues.Preserve;
                }
                else
                {
                    secondText.Space = null;
                }
                if (!(char.IsWhiteSpace(text.Text[0]) || char.IsWhiteSpace(text.Text[text.Text.Length - 1])))
                {
                    text.Space = null;
                }
                BookmarkStart start = new BookmarkStart
                {
                    Id = (nextBookmark++).ToString(CultureInfo.InvariantCulture),
                    Name = "_Toc" + (nextName++).ToString(CultureInfo.InvariantCulture)
                };
                run.InsertAfterSelf(second);
                second.InsertBeforeSelf(start);
                paragraph.AppendChild(new BookmarkEnd { Id = start.Id.Value });
                return start.Name.Value;
            }
            seen += length;
        }
        return EnsureHeadingBookmark(paragraph, ref nextBookmark, ref nextName);
    }

    //The name of the _Toc bookmark of a TC field. Word puts it around the entry inside the code of the field, which is then written in three parts.
    internal static string EnsureTcBookmark(TocItem item, ref uint nextBookmark, ref long nextName)
    {
        Paragraph paragraph = item.Paragraph;
        List<OpenXmlElement> children = paragraph.ChildElements.ToList();
        int begin = children.IndexOf(item.TcBegin);
        int end = children.IndexOf(item.TcEnd);
        BookmarkStart existing = children.Skip(begin).Take(end - begin + 1).OfType<BookmarkStart>().FirstOrDefault(b => b.Name?.Value != null && b.Name.Value.StartsWith("_Toc", StringComparison.Ordinal));
        if (existing != null)
        {
            return existing.Name.Value;
        }

        foreach (Run code in children.Skip(begin + 1).Take(end - begin - 1).OfType<Run>().Where(r => r.Elements<FieldCode>().Any()).ToList())
        {
            code.Remove();
        }
        BookmarkStart start = new BookmarkStart
        {
            Id = (nextBookmark++).ToString(CultureInfo.InvariantCulture),
            Name = "_Toc" + (nextName++).ToString(CultureInfo.InvariantCulture)
        };
        BookmarkEnd stop = new BookmarkEnd { Id = start.Id.Value };
        OpenXmlElement last = item.TcBegin;
        foreach (OpenXmlElement part in new OpenXmlElement[] { CodeRun(" TC \""), start, CodeRun(item.TcEntry), stop, CodeRun("\"" + item.TcRest) })
        {
            last.InsertAfterSelf(part);
            last = part;
        }
        return start.Name.Value;
    }
}
