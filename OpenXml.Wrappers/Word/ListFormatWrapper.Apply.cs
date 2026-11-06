using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The lists Word starts a list from: its default bullet, number and outline list templates, kept as embedded definitions
/// (<c>Templates/ListDefaults.xml</c>, written by <c>tools/WordReference</c>).
/// </summary>
internal static class ListDefaults
{
    internal enum Kind
    {
        Bullet,
        Number,
        Outline
    }

    private static readonly Random Ids = new Random();
    private static readonly object IdLock = new object();

    private static Numbering ReadTemplate()
    {
        using (Stream stream = NewDocument.OpenTemplate("ListDefaults.xml"))
        using (StreamReader reader = new StreamReader(stream))
        {
            //Without the declaration of the document, which an element cannot start with.
            string xml = reader.ReadToEnd();
            return new Numbering(xml.Substring(xml.IndexOf("?>", StringComparison.Ordinal) + 2));
        }
    }

    /// <summary>The kind of default list a definition is.</summary>
    internal static Kind Classify(AbstractNum template)
    {
        if (template.Elements<Level>().FirstOrDefault()?.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
        {
            return Kind.Bullet;
        }
        return template.MultiLevelType?.Val?.Value == MultiLevelValues.Multilevel ? Kind.Outline : Kind.Number;
    }

    internal static int NewRandom()
    {
        lock (IdLock)
        {
            return Ids.Next(0x10000000, int.MaxValue);
        }
    }

    internal static string NewNsid()
    {
        return NewRandom().ToString("X8", CultureInfo.InvariantCulture);
    }

    /// <summary>The numbering of the document, which a document without lists does not have yet.</summary>
    internal static Numbering GetOrCreateNumbering(MainDocumentPart main)
    {
        NumberingDefinitionsPart part = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
        if (part.Numbering == null)
        {
            Numbering root = ReadTemplate();
            root.RemoveAllChildren();
            part.Numbering = root;
        }
        return part.Numbering;
    }

    /// <summary>A copy of a definition, with an id and nsid of its own, added to the numbering.</summary>
    internal static AbstractNum AddCopy(Numbering numbering, AbstractNum source)
    {
        AbstractNum copy = (AbstractNum)source.CloneNode(true);
        int id = numbering.Elements<AbstractNum>().Select(a => a.AbstractNumberId?.Value ?? -1).DefaultIfEmpty(-1).Max() + 1;
        copy.AbstractNumberId = id;
        copy.Nsid = new Nsid() { Val = NewNsid() };
        NumberingInstance firstNum = numbering.Elements<NumberingInstance>().FirstOrDefault();
        if (firstNum != null)
        {
            firstNum.InsertBeforeSelf(copy);
        }
        else
        {
            numbering.AppendChild(copy);
        }
        return copy;
    }

    internal static AbstractNum AddDefault(Numbering numbering, Kind kind)
    {
        AbstractNum source = ReadTemplate().Elements<AbstractNum>().First(a => Classify(a) == kind);
        return AddCopy(numbering, source);
    }

    /// <summary>A new list that uses a definition.</summary>
    internal static NumberingInstance AddNum(Numbering numbering, AbstractNum template)
    {
        int id = numbering.Elements<NumberingInstance>().Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
        NumberingInstance num = new NumberingInstance(new AbstractNumId() { Val = template.AbstractNumberId.Value }) { NumberID = id };
        num.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", NewRandom().ToString(CultureInfo.InvariantCulture)));
        numbering.AppendChild(num);
        return num;
    }
}

public partial class ListFormatWrapper
{
    private MainDocumentPart GetMain()
    {
        return StyleHelper.GetMainDocumentPart(GetParagraphs().First());
    }

    //The numbered paragraph before the first paragraph of the range, if there is one.
    private Paragraph GetPreviousParagraph(Paragraph first)
    {
        List<Paragraph> all = StoryMap.FindRoot(first).Descendants<Paragraph>().ToList();
        int at = all.IndexOf(first);
        return at > 0 ? all[at - 1] : null;
    }

    private void ApplyDefault(ListDefaults.Kind kind)
    {
        List<Paragraph> paragraphs = GetParagraphs();
        if (paragraphs.Count == 0)
        {
            return;
        }

        MainDocumentPart main = GetMain();
        BuiltInStyles.GetOrCreate(main, "ListParagraph");
        Numbering numbering = ListDefaults.GetOrCreateNumbering(main);

        NumberingInstance num = null;
        if (kind == ListDefaults.Kind.Bullet)
        {
            //Bullets do not count, so the lists of one kind of bullet are one list.
            num = numbering.Elements<NumberingInstance>().FirstOrDefault(n => FindTemplate(numbering, n) is AbstractNum a && ListDefaults.Classify(a) == kind);
        }
        else
        {
            //A numbered list goes on from the paragraph before it when that is numbered the same way, and starts again otherwise.
            Paragraph previous = GetPreviousParagraph(paragraphs[0]);
            if (previous != null && ListNumbering.TryGetNumbering(previous, out int previousId, out _))
            {
                NumberingInstance previousNum = numbering.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == previousId);
                AbstractNum previousTemplate = previousNum == null ? null : FindTemplate(numbering, previousNum);
                if (previousTemplate != null && ListDefaults.Classify(previousTemplate) == kind)
                {
                    num = previousNum;
                }
            }
        }
        if (num == null)
        {
            num = ListDefaults.AddNum(numbering, ListDefaults.AddDefault(numbering, kind));
        }

        int hanging = LevelHanging(FindTemplate(numbering, num), 0);
        foreach (Paragraph paragraph in paragraphs)
        {
            Attach(paragraph, num, 0, hanging);
        }
    }

    private static AbstractNum FindTemplate(Numbering numbering, NumberingInstance num)
    {
        int? id = num.AbstractNumId?.Val?.Value;
        return id == null ? null : numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == id.Value);
    }

    private static int LevelHanging(AbstractNum template, int level)
    {
        Indentation indentation = ListNumbering.FindLevel(template, level)?.PreviousParagraphProperties?.Indentation;
        return indentation?.Hanging?.Value != null && int.TryParse(indentation.Hanging.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int hanging) ? hanging : 0;
    }

    //Puts a paragraph in a list: the list paragraph style (unless it has a style of its own), the list and level, and a left indent it had
    //moved on by the space for the number.
    private static void Attach(Paragraph paragraph, NumberingInstance num, int level, int hanging)
    {
        ParagraphProperties properties = paragraph.GetOrAddParagraphProperties();
        string style = properties.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(style) || style == "Normal")
        {
            properties.ParagraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };
        }
        properties.NumberingProperties = new NumberingProperties(new NumberingLevelReference() { Val = level }, new NumberingId() { Val = num.NumberID.Value });

        Indentation indentation = properties.Indentation;
        string left = indentation?.Left?.Value ?? indentation?.Start?.Value;
        if (left != null && int.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips))
        {
            indentation.Left = (twips + hanging).ToString(CultureInfo.InvariantCulture);
            indentation.Start = null;
        }
    }

    /// <summary>Makes the paragraphs the items of a bulleted list with Word's default bullet, as <c>ListFormat.ApplyBulletDefault</c> in Word.</summary>
    public void ApplyBulletDefault()
    {
        ApplyDefault(ListDefaults.Kind.Bullet);
    }

    /// <summary>
    /// Makes the paragraphs the items of a numbered list (1., 2., 3.), as <c>ListFormat.ApplyNumberDefault</c> in Word. The list goes on from
    /// the paragraph before the range when that is numbered the same way.
    /// </summary>
    public void ApplyNumberDefault()
    {
        ApplyDefault(ListDefaults.Kind.Number);
    }

    /// <summary>
    /// Makes the paragraphs the items of an outline list (1), a), i) and so on), as <c>ListFormat.ApplyOutlineNumberDefault</c> in Word.
    /// </summary>
    public void ApplyOutlineNumberDefault()
    {
        ApplyDefault(ListDefaults.Kind.Outline);
    }

    /// <summary>
    /// Makes the paragraphs the items of a list that uses a template, as <c>ListFormat.ApplyListTemplate</c> in Word. With
    /// <paramref name="continuePreviousList"/> the paragraphs go on from the list before the range when that uses the template; otherwise they
    /// start a list of their own, with a copy of the template if it is in use.
    /// </summary>
    /// <param name="template">The template of the document to use.</param>
    /// <param name="continuePreviousList">True to carry on the list before the range.</param>
    public void ApplyListTemplate(ListTemplateWrapper template, bool continuePreviousList = false)
    {
        Guard.ThrowIfNull(template);
        List<Paragraph> paragraphs = GetParagraphs();
        if (paragraphs.Count == 0)
        {
            return;
        }

        MainDocumentPart main = GetMain();
        BuiltInStyles.GetOrCreate(main, "ListParagraph");
        Numbering numbering = ListDefaults.GetOrCreateNumbering(main);
        if (!ReferenceEquals(template.AbstractNum.Parent, numbering))
        {
            throw new ArgumentException("The template is not one of this document.", nameof(template));
        }

        NumberingInstance num = null;
        if (continuePreviousList)
        {
            Paragraph previous = GetPreviousParagraph(paragraphs[0]);
            if (previous != null && ListNumbering.TryGetNumbering(previous, out int previousId, out _))
            {
                NumberingInstance previousNum = numbering.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == previousId);
                if (previousNum != null && FindTemplate(numbering, previousNum) == template.AbstractNum)
                {
                    num = previousNum;
                }
            }
        }
        if (num == null)
        {
            bool inUse = numbering.Elements<NumberingInstance>().Any(n => FindTemplate(numbering, n) == template.AbstractNum);
            num = ListDefaults.AddNum(numbering, inUse ? ListDefaults.AddCopy(numbering, template.AbstractNum) : template.AbstractNum);
        }

        int hanging = LevelHanging(FindTemplate(numbering, num), 0);
        foreach (Paragraph paragraph in paragraphs)
        {
            Attach(paragraph, num, 0, hanging);
        }
    }

    /// <summary>
    /// Takes the paragraphs out of their lists, as <c>ListFormat.RemoveNumbers</c> in Word. The list paragraph style goes with the numbering.
    /// </summary>
    public void RemoveNumbers()
    {
        foreach (ListNumbering.Item item in GetItems())
        {
            ParagraphProperties properties = item.Paragraph.ParagraphProperties;
            properties.NumberingProperties = null;
            if (properties.ParagraphStyleId?.Val?.Value == "ListParagraph")
            {
                properties.ParagraphStyleId = null;
            }
            if (!properties.HasChildren)
            {
                properties.Remove();
            }
        }
    }

    private void MoveLevel(int change)
    {
        foreach (ListNumbering.Item item in GetItems())
        {
            int levels = Math.Min(9, item.Template.Elements<Level>().Count());
            int level = Math.Max(0, Math.Min(levels - 1, item.Level + change));
            SetLevel(item.Paragraph, level);
        }
    }

    private static void SetLevel(Paragraph paragraph, int level)
    {
        NumberingProperties numbering = paragraph.ParagraphProperties?.NumberingProperties;
        if (numbering != null)
        {
            numbering.NumberingLevelReference = new NumberingLevelReference() { Val = level };
        }
    }

    /// <summary>Moves the paragraphs one level deeper in their lists, as <c>ListFormat.ListIndent</c> in Word; the deepest level stays.</summary>
    public void ListIndent()
    {
        MoveLevel(1);
    }

    /// <summary>Moves the paragraphs one level up in their lists, as <c>ListFormat.ListOutdent</c> in Word; the first level stays.</summary>
    public void ListOutdent()
    {
        MoveLevel(-1);
    }

    private void SetListLevelNumber(int level)
    {
        foreach (ListNumbering.Item item in GetItems())
        {
            if (level < 1 || level > Math.Min(9, item.Template.Elements<Level>().Count()))
            {
                throw new ArgumentOutOfRangeException(nameof(level), "The list has no such level.");
            }
            SetLevel(item.Paragraph, level - 1);
        }
    }

    /// <summary>
    /// Turns the numbers or bullets of the paragraphs into text, as <c>ListFormat.ConvertNumbersToText</c> in Word: the paragraphs keep their
    /// indents and what the number said, a tab or space after it, and are no longer in a list.
    /// </summary>
    public void ConvertNumbersToText()
    {
        foreach (ListNumbering.Item item in GetItems())
        {
            ConvertToText(item);
        }
    }

    private static void ConvertToText(ListNumbering.Item item)
    {
        Paragraph paragraph = item.Paragraph;
        Level level = ListNumbering.FindLevel(item.Template, item.Level);
        RunProperties numberProperties = new RunProperties();
        foreach (OpenXmlElement child in level?.NumberingSymbolRunProperties?.ChildElements ?? Enumerable.Empty<OpenXmlElement>())
        {
            OpenXmlElement copy = child.CloneNode(true);
            //The hint says which font a character is set in, and Word leaves it out of the text it makes.
            if (copy is RunFonts fonts)
            {
                fonts.Hint = null;
            }
            numberProperties.AppendChild(copy);
        }
        bool plain = !numberProperties.HasChildren;
        LevelSuffixValues? suffix = level?.LevelSuffix?.Val?.Value;

        Run first = paragraph.Elements<Run>().FirstOrDefault();
        Run number = new Run();
        if (!plain)
        {
            number.RunProperties = (RunProperties)numberProperties.CloneNode(true);
        }
        number.AppendChild(OpenXmlExtensions.NewText(suffix == LevelSuffixValues.Space ? item.Text + " " : item.Text));
        if (first != null)
        {
            first.InsertBeforeSelf(number);
        }
        else
        {
            paragraph.AppendChild(number);
        }

        if (suffix == null || suffix == LevelSuffixValues.Tab)
        {
            if (plain && first != null && first.RunProperties == null)
            {
                first.PrependChild(new TabChar());
            }
            else
            {
                Run tab = new Run(new TabChar());
                if (!plain)
                {
                    tab.RunProperties = (RunProperties)numberProperties.CloneNode(true);
                }
                number.InsertAfterSelf(tab);
            }
        }

        ParagraphProperties properties = paragraph.GetOrAddParagraphProperties();
        properties.NumberingProperties = null;
        if (properties.ParagraphStyleId?.Val?.Value == "ListParagraph")
        {
            properties.ParagraphStyleId = null;
        }
        Indentation levelIndentation = level?.PreviousParagraphProperties?.Indentation;
        if (levelIndentation != null)
        {
            properties.Indentation = (Indentation)levelIndentation.CloneNode(true);
        }
    }
}
