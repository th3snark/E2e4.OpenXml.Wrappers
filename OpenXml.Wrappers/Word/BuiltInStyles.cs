using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The styles that come with Word. A document only holds the ones that have been used, but in Word they can all be asked for by name, so
/// the definitions are kept here and added to a document when one is asked for or applied.
/// </summary>
internal static class BuiltInStyles
{
    private sealed class Definition
    {
        public BuiltInStyle? Id;
        public string Name;
        public string StoredName;
        public string StyleId;
        public StyleValues Type = StyleValues.Paragraph;
        public string BasedOn;
        public string Next;
        public bool IsDefault;
        public bool AutoRedefine;
        public int UiPriority;
        public bool SemiHidden;
        public bool UnhideWhenUsed;
        public bool QuickFormat;
        public Func<StyleParagraphProperties> Paragraph;
        public Func<StyleRunProperties> Run;
        public Func<StyleTableProperties> Table;
    }

    private static readonly List<Definition> Definitions = CreateDefinitions();
    private static readonly Dictionary<BuiltInStyle, Definition> ById = Definitions.Where(d => d.Id.HasValue).ToDictionary(d => d.Id.Value);

    /// <summary>The name Word shows for the built-in style that is stored under the name, or null if it is not one that is known.</summary>
    public static string GetDisplayName(string storedName)
    {
        return storedName == null ? null : Find(storedName)?.Name;
    }

    public static bool IsBuiltInName(string name)
    {
        return name != null && Find(name) != null;
    }

    /// <summary>The name or id of a built-in style that has a definition.</summary>
    public static string GetName(BuiltInStyle style)
    {
        if (!ById.TryGetValue(style, out Definition definition))
        {
            throw new ArgumentOutOfRangeException(nameof(style));
        }
        return definition.Name;
    }

    /// <summary>
    /// Finds the style in the document and, if the document does not have it but it is a built-in one, adds it. Returns null when there is no
    /// such style.
    /// </summary>
    public static Style GetOrCreate(MainDocumentPart mainDocumentPart, string nameOrId)
    {
        Style existing = StyleHelper.FindStyle(mainDocumentPart?.StyleDefinitionsPart?.Styles, nameOrId);
        if (existing != null || mainDocumentPart == null)
        {
            return existing;
        }

        Definition definition = Find(nameOrId);
        return definition == null ? null : Create(mainDocumentPart, definition);
    }

    private static Definition Find(string nameOrId)
    {
        return Definitions.FirstOrDefault(d => string.Equals(d.Name, nameOrId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(d.StoredName, nameOrId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(d.StyleId, nameOrId, StringComparison.OrdinalIgnoreCase));
    }

    private static Style Create(MainDocumentPart mainDocumentPart, Definition definition)
    {
        StyleDefinitionsPart part = mainDocumentPart.StyleDefinitionsPart ?? mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        if (part.Styles == null)
        {
            part.Styles = new Styles();
        }

        Style existing = part.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == definition.StyleId);
        if (existing != null)
        {
            return existing;
        }

        //The styles that this one is based on and followed by come first, so that the document never refers to a style it does not have.
        foreach (string dependency in new[] { definition.BasedOn, definition.Next })
        {
            Definition other = dependency == null || dependency == definition.StyleId ? null : Find(dependency);
            if (other != null)
            {
                Create(mainDocumentPart, other);
            }
        }

        Style style = new Style(new StyleName() { Val = definition.StoredName }) { Type = definition.Type, StyleId = definition.StyleId };
        if (definition.IsDefault)
        {
            style.Default = true;
        }
        if (definition.BasedOn != null)
        {
            style.AppendChild(new BasedOn() { Val = definition.BasedOn });
        }
        if (definition.Next != null)
        {
            style.AppendChild(new NextParagraphStyle() { Val = definition.Next });
        }
        if (definition.AutoRedefine)
        {
            style.AppendChild(new AutoRedefine());
        }
        if (definition.UiPriority > 0)
        {
            style.AppendChild(new UIPriority() { Val = definition.UiPriority });
        }
        if (definition.SemiHidden)
        {
            style.AppendChild(new SemiHidden());
        }
        if (definition.UnhideWhenUsed)
        {
            style.AppendChild(new UnhideWhenUsed());
        }
        if (definition.QuickFormat)
        {
            style.AppendChild(new PrimaryStyle());
        }
        if (definition.Paragraph != null)
        {
            style.AppendChild(definition.Paragraph());
        }
        if (definition.Run != null)
        {
            style.AppendChild(definition.Run());
        }
        if (definition.Table != null)
        {
            style.AppendChild(definition.Table());
        }

        //The default styles go first, as Word writes them.
        if (definition.IsDefault)
        {
            part.Styles.InsertAt(style, part.Styles.Elements<DocDefaults>().Any() ? 1 : 0);
        }
        else
        {
            part.Styles.AppendChild(style);
        }
        return style;
    }

    #region Definitions

    private static StyleRunProperties Run(params OpenXmlElement[] properties)
    {
        return new StyleRunProperties(properties);
    }

    private static StyleParagraphProperties Paragraph(params OpenXmlElement[] properties)
    {
        return new StyleParagraphProperties(properties);
    }

    private static SpacingBetweenLines Spacing(int? before, int? after, int? line)
    {
        SpacingBetweenLines spacing = new SpacingBetweenLines();
        if (before.HasValue)
        {
            spacing.Before = before.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (after.HasValue)
        {
            spacing.After = after.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (line.HasValue)
        {
            spacing.Line = line.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            spacing.LineRule = LineSpacingRuleValues.Auto;
        }
        return spacing;
    }

    private static RunFonts Fonts(string name)
    {
        return new RunFonts() { Ascii = name, HighAnsi = name };
    }

    private static FontSize Size(int halfPoints)
    {
        return new FontSize() { Val = halfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture) };
    }

    private static List<Definition> CreateDefinitions()
    {
        List<Definition> list = new List<Definition>();

        Definition Add(BuiltInStyle? id, string name, string storedName, string styleId, StyleValues type, string basedOn, string next, int priority, bool quick,
            bool semiHidden = false, bool unhide = false, Func<StyleParagraphProperties> paragraph = null, Func<StyleRunProperties> run = null)
        {
            Definition definition = new Definition()
            {
                Id = id,
                Name = name,
                StoredName = storedName,
                StyleId = styleId,
                Type = type,
                BasedOn = basedOn,
                Next = next,
                UiPriority = priority,
                QuickFormat = quick,
                SemiHidden = semiHidden,
                UnhideWhenUsed = unhide,
                Paragraph = paragraph,
                Run = run
            };
            list.Add(definition);
            return definition;
        }

        Definition normal = Add(BuiltInStyle.Normal, "Normal", "Normal", "Normal", StyleValues.Paragraph, null, null, 0, true);
        normal.IsDefault = true;
        Definition defaultFont = Add(BuiltInStyle.DefaultParagraphFont, "Default Paragraph Font", "Default Paragraph Font", "DefaultParagraphFont", StyleValues.Character, null, null, 1, false, true, true);
        defaultFont.IsDefault = true;
        Definition normalTable = Add(BuiltInStyle.NormalTable, "Normal Table", "Normal Table", "TableNormal", StyleValues.Table, null, null, 99, false, true, true);
        normalTable.IsDefault = true;
        normalTable.Table = () => new StyleTableProperties(
            new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa },
            new TableCellMarginDefault(
                new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa },
                new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa }));
        Definition tableGrid = Add(null, "Table Grid", "Table Grid", "TableGrid", StyleValues.Table, "TableNormal", null, 39, false,
            paragraph: () => Paragraph(Spacing(null, 0, 240)));
        tableGrid.Table = () => new StyleTableProperties(new TableBorders(
            new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 },
            new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 },
            new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 },
            new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 },
            new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 },
            new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 }));
        Definition noList = Add(null, "No List", "No List", "NoList", StyleValues.Numbering, null, null, 99, false, true, true);
        noList.IsDefault = true;

        //Headings
        BuiltInStyle[] headings = { BuiltInStyle.Heading1, BuiltInStyle.Heading2, BuiltInStyle.Heading3, BuiltInStyle.Heading4, BuiltInStyle.Heading5, BuiltInStyle.Heading6, BuiltInStyle.Heading7, BuiltInStyle.Heading8, BuiltInStyle.Heading9 };
        string[] headingColors = { "2F5496", "2F5496", "1F3763", "2F5496", "2F5496", "1F3763", "1F3763", "272727", "272727" };
        int[] headingSizes = { 32, 26, 24, 0, 0, 0, 0, 21, 21 };
        for (int i = 0; i < 9; i++)
        {
            int level = i;
            Add(headings[i], "Heading " + (i + 1), "heading " + (i + 1), "Heading" + (i + 1), StyleValues.Paragraph, "Normal", "Normal", 9, true, false, i > 0,
                () => Paragraph(new KeepNext(), new KeepLines(), Spacing(level == 0 ? 240 : 40, 0, null), new OutlineLevel() { Val = level }),
                () =>
                {
                    StyleRunProperties run = Run(Fonts(level < 3 ? "Calibri Light" : "Calibri"));
                    if (level == 3 || level == 6 || level == 8)
                    {
                        run.AppendChild(new Italic());
                        run.AppendChild(new ItalicComplexScript());
                    }
                    run.AppendChild(new Color() { Val = headingColors[level] });
                    if (headingSizes[level] > 0)
                    {
                        run.AppendChild(Size(headingSizes[level]));
                        run.AppendChild(new FontSizeComplexScript() { Val = headingSizes[level].ToString(System.Globalization.CultureInfo.InvariantCulture) });
                    }
                    return run;
                });
        }

        Add(BuiltInStyle.Title, "Title", "Title", "Title", StyleValues.Paragraph, "Normal", "Normal", 10, true, false, false,
            () => Paragraph(Spacing(null, 0, 240), new ContextualSpacing()),
            () => Run(Fonts("Calibri Light"), new Spacing() { Val = -10 }, new Kern() { Val = 28 }, Size(56), new FontSizeComplexScript() { Val = "56" }));
        Add(BuiltInStyle.Subtitle, "Subtitle", "Subtitle", "Subtitle", StyleValues.Paragraph, "Normal", "Normal", 11, true, false, false,
            () => Paragraph(Spacing(null, 160, null)),
            () => Run(new Color() { Val = "5A5A5A" }, new Spacing() { Val = 15 }));
        Add(BuiltInStyle.Quote, "Quote", "Quote", "Quote", StyleValues.Paragraph, "Normal", "Normal", 29, true, false, false,
            () => Paragraph(Spacing(200, 160, null), new Indentation() { Left = "864", Right = "864" }, new Justification() { Val = JustificationValues.Center }),
            () => Run(new Italic(), new ItalicComplexScript(), new Color() { Val = "404040" }));
        Add(BuiltInStyle.IntenseQuote, "Intense Quote", "Intense Quote", "IntenseQuote", StyleValues.Paragraph, "Normal", "Normal", 30, true, false, false,
            () => Paragraph(
                new ParagraphBorders(
                    new TopBorder() { Val = BorderValues.Single, Color = "4472C4", Size = 4, Space = 10 },
                    new BottomBorder() { Val = BorderValues.Single, Color = "4472C4", Size = 4, Space = 10 }),
                Spacing(360, 360, null),
                new Indentation() { Left = "864", Right = "864" },
                new Justification() { Val = JustificationValues.Center }),
            () => Run(new Italic(), new ItalicComplexScript(), new Color() { Val = "4472C4" }));
        Add(BuiltInStyle.ListParagraph, "List Paragraph", "List Paragraph", "ListParagraph", StyleValues.Paragraph, "Normal", null, 34, true, false, false,
            () => Paragraph(new Indentation() { Left = "720" }, new ContextualSpacing()));
        Add(BuiltInStyle.Caption, "Caption", "caption", "Caption", StyleValues.Paragraph, "Normal", "Normal", 35, true, false, true,
            () => Paragraph(Spacing(null, 200, 240)),
            () => Run(new Italic(), new ItalicComplexScript(), new Color() { Val = "44546A" }, Size(18), new FontSizeComplexScript() { Val = "18" }));
        Add(BuiltInStyle.BodyText, "Body Text", "Body Text", "BodyText", StyleValues.Paragraph, "Normal", null, 0, false, false, false,
            () => Paragraph(Spacing(null, 120, null)));
        Add(BuiltInStyle.BodyText2, "Body Text 2", "Body Text 2", "BodyText2", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, 120, 480)));
        Add(BuiltInStyle.BodyText3, "Body Text 3", "Body Text 3", "BodyText3", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, 120, null)),
            () => Run(Size(16), new FontSizeComplexScript() { Val = "16" }));
        Add(BuiltInStyle.BodyTextIndent, "Body Text Indent", "Body Text Indent", "BodyTextIndent", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, 120, null), new Indentation() { Left = "283" }));
        Add(BuiltInStyle.PlainText, "Plain Text", "Plain Text", "PlainText", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, 0, 240)),
            () => Run(new RunFonts() { Ascii = "Consolas", HighAnsi = "Consolas" }, Size(21), new FontSizeComplexScript() { Val = "21" }));
        Add(BuiltInStyle.Closing, "Closing", "Closing", "Closing", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, 0, 240), new Indentation() { Left = "4320" }));
        Add(BuiltInStyle.Signature, "Signature", "Signature", "Signature", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, 0, 240), new Indentation() { Left = "4320" }));
        Add(BuiltInStyle.Salutation, "Salutation", "Salutation", "Salutation", StyleValues.Paragraph, "Normal", "Normal", 99, false, false, true);
        Add(BuiltInStyle.Date, "Date", "Date", "Date", StyleValues.Paragraph, "Normal", "Normal", 99, false, false, true);

        Definition header = Add(BuiltInStyle.Header, "Header", "header", "Header", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(new Tabs(new TabStop() { Val = TabStopValues.Center, Position = 4680 }, new TabStop() { Val = TabStopValues.Right, Position = 9360 }), Spacing(null, 0, 240)));
        header.AutoRedefine = false;
        Add(BuiltInStyle.Footer, "Footer", "footer", "Footer", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(new Tabs(new TabStop() { Val = TabStopValues.Center, Position = 4680 }, new TabStop() { Val = TabStopValues.Right, Position = 9360 }), Spacing(null, 0, 240)));
        Add(BuiltInStyle.FootnoteText, "Footnote Text", "footnote text", "FootnoteText", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, 0, 240)),
            () => Run(Size(20), new FontSizeComplexScript() { Val = "20" }));
        Add(BuiltInStyle.EndnoteText, "Endnote Text", "endnote text", "EndnoteText", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, 0, 240)),
            () => Run(Size(20), new FontSizeComplexScript() { Val = "20" }));
        Add(BuiltInStyle.CommentText, "Comment Text", "annotation text", "CommentText", StyleValues.Paragraph, "Normal", null, 99, false, false, true,
            () => Paragraph(Spacing(null, null, 240)),
            () => Run(Size(20), new FontSizeComplexScript() { Val = "20" }));

        for (int i = 1; i <= 9; i++)
        {
            int level = i;
            Definition toc = Add((BuiltInStyle)(-19 - i), "TOC " + i, "toc " + i, "TOC" + i, StyleValues.Paragraph, "Normal", "Normal", 39, false, false, true,
                () =>
                {
                    StyleParagraphProperties properties = Paragraph(Spacing(null, 100, null));
                    if (level > 1)
                    {
                        properties.AppendChild(new Indentation() { Left = ((level - 1) * 220).ToString(System.Globalization.CultureInfo.InvariantCulture) });
                    }
                    return properties;
                });
            toc.AutoRedefine = true;
        }

        //Character styles
        Add(BuiltInStyle.Hyperlink, "Hyperlink", "Hyperlink", "Hyperlink", StyleValues.Character, null, null, 99, false, false, true,
            run: () => Run(new Color() { Val = "0563C1" }, new Underline() { Val = UnderlineValues.Single }));
        Add(BuiltInStyle.HyperlinkFollowed, "Followed Hyperlink", "FollowedHyperlink", "FollowedHyperlink", StyleValues.Character, null, null, 99, false, false, true,
            run: () => Run(new Color() { Val = "954F72" }, new Underline() { Val = UnderlineValues.Single }));
        Add(BuiltInStyle.Strong, "Strong", "Strong", "Strong", StyleValues.Character, null, null, 22, true,
            run: () => Run(new Bold(), new BoldComplexScript()));
        Add(BuiltInStyle.Emphasis, "Emphasis", "Emphasis", "Emphasis", StyleValues.Character, null, null, 20, true,
            run: () => Run(new Italic(), new ItalicComplexScript()));
        Add(BuiltInStyle.SubtleEmphasis, "Subtle Emphasis", "Subtle Emphasis", "SubtleEmphasis", StyleValues.Character, null, null, 19, true,
            run: () => Run(new Italic(), new ItalicComplexScript(), new Color() { Val = "404040" }));
        Add(BuiltInStyle.IntenseEmphasis, "Intense Emphasis", "Intense Emphasis", "IntenseEmphasis", StyleValues.Character, null, null, 21, true,
            run: () => Run(new Italic(), new ItalicComplexScript(), new Color() { Val = "4472C4" }));
        Add(BuiltInStyle.SubtleReference, "Subtle Reference", "Subtle Reference", "SubtleReference", StyleValues.Character, null, null, 31, true,
            run: () => Run(new SmallCaps(), new Color() { Val = "5A5A5A" }));
        Add(BuiltInStyle.IntenseReference, "Intense Reference", "Intense Reference", "IntenseReference", StyleValues.Character, null, null, 32, true,
            run: () => Run(new Bold(), new BoldComplexScript(), new SmallCaps(), new Color() { Val = "4472C4" }, new Spacing() { Val = 5 }));
        Add(BuiltInStyle.BookTitle, "Book Title", "Book Title", "BookTitle", StyleValues.Character, null, null, 33, true,
            run: () => Run(new Bold(), new BoldComplexScript(), new Italic(), new ItalicComplexScript(), new Spacing() { Val = 5 }));
        Add(BuiltInStyle.FootnoteReference, "Footnote Reference", "footnote reference", "FootnoteReference", StyleValues.Character, null, null, 99, false, false, true,
            run: () => Run(new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript }));
        Add(BuiltInStyle.EndnoteReference, "Endnote Reference", "endnote reference", "EndnoteReference", StyleValues.Character, null, null, 99, false, false, true,
            run: () => Run(new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript }));
        Add(BuiltInStyle.CommentReference, "Comment Reference", "annotation reference", "CommentReference", StyleValues.Character, null, null, 99, false, true, true,
            run: () => Run(Size(16), new FontSizeComplexScript() { Val = "16" }));
        Add(BuiltInStyle.PageNumber, "Page Number", "page number", "PageNumber", StyleValues.Character, null, null, 0, false);
        Add(BuiltInStyle.LineNumber, "Line Number", "line number", "LineNumber", StyleValues.Character, null, null, 0, false);

        return list;
    }

    #endregion
}
