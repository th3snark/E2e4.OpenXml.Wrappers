using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// How the numbers of a list level are written, as <c>WdListNumberStyle</c> in Word (with the same values, for the styles the library knows).
/// </summary>
public enum ListNumberStyle
{
    /// <summary>1, 2, 3.</summary>
    Arabic = 0,

    /// <summary>I, II, III.</summary>
    UppercaseRoman = 1,

    /// <summary>i, ii, iii.</summary>
    LowercaseRoman = 2,

    /// <summary>A, B, C.</summary>
    UppercaseLetter = 3,

    /// <summary>a, b, c.</summary>
    LowercaseLetter = 4,

    /// <summary>1st, 2nd, 3rd.</summary>
    Ordinal = 5,

    /// <summary>One, Two, Three.</summary>
    CardinalText = 6,

    /// <summary>First, Second, Third.</summary>
    OrdinalText = 7,

    /// <summary>01, 02, 03.</summary>
    ArabicLZ = 22,

    /// <summary>A bullet.</summary>
    Bullet = 23,

    /// <summary>No number.</summary>
    None = 255
}

/// <summary>
/// What follows the number of a list item, as <c>WdTrailingCharacter</c> in Word (with the same values).
/// </summary>
public enum TrailingCharacter
{
    /// <summary>A tab.</summary>
    Tab = 0,

    /// <summary>A space.</summary>
    Space = 1,

    /// <summary>Nothing.</summary>
    None = 2
}

/// <summary>
/// How the number of a list item is lined up in the space for it, as <c>WdListLevelAlignment</c> in Word (with the same values).
/// </summary>
public enum ListLevelAlignment
{
    /// <summary>Left.</summary>
    Left = 0,

    /// <summary>Centred.</summary>
    Center = 1,

    /// <summary>Right.</summary>
    Right = 2
}

/// <summary>
/// What kind of list the paragraphs of a range are in, as <c>WdListType</c> in Word (with the same values).
/// </summary>
public enum ListType
{
    /// <summary>No list.</summary>
    NoNumbering = 0,

    /// <summary>Numbers only (a numbered list that does not use the outline levels).</summary>
    ListNumberOnly = 1,

    /// <summary>Bullets.</summary>
    Bullet = 2,

    /// <summary>Numbers in a simple list.</summary>
    SimpleNumbering = 3,

    /// <summary>Numbers in an outline of several levels.</summary>
    OutlineNumbering = 4,

    /// <summary>Different kinds of list.</summary>
    MixedNumbering = 5,

    /// <summary>Pictures as bullets.</summary>
    PictureBullet = 6
}

/// <summary>
/// A list template, as <c>ListTemplate</c> in Word: the definition of how the items of a list are numbered (a <c>w:abstractNum</c>).
/// </summary>
public partial class ListTemplateWrapper : IEquatable<ListTemplateWrapper>
{
    private readonly AbstractNum _abstractNum;

    internal ListTemplateWrapper(AbstractNum abstractNum)
    {
        _abstractNum = abstractNum;
    }

    internal AbstractNum AbstractNum
    {
        get { return _abstractNum; }
    }

    /// <summary>The name of the template, as <c>ListTemplate.Name</c> in Word; empty if it has none.</summary>
    public string Name
    {
        get { return _abstractNum.AbstractNumDefinitionName?.Val?.Value ?? string.Empty; }
    }

    /// <summary>True if the template has the levels of an outline, as <c>ListTemplate.OutlineNumbered</c> in Word.</summary>
    public bool OutlineNumbered
    {
        get { return _abstractNum.MultiLevelType?.Val?.Value == MultiLevelValues.Multilevel; }
    }

    /// <summary>The levels of the template, as <c>ListTemplate.ListLevels</c> in Word.</summary>
    public ListLevelWrapperCollection ListLevels
    {
        get { return new ListLevelWrapperCollection(this); }
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as ListTemplateWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(ListTemplateWrapper other)
    {
        return !(other is null) && ReferenceEquals(_abstractNum, other._abstractNum);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _abstractNum.GetHashCode();
    }

    public static bool operator ==(ListTemplateWrapper left, ListTemplateWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(ListTemplateWrapper left, ListTemplateWrapper right)
    {
        return !(left == right);
    }
}

/// <summary>
/// The list templates of a document, as <c>ListTemplates</c> in Word.
/// </summary>
public partial class ListTemplateWrapperCollection : IEnumerable<ListTemplateWrapper>
{
    private readonly WordprocessingDocumentWrapper _document;

    internal ListTemplateWrapperCollection(WordprocessingDocumentWrapper document)
    {
        _document = document;
    }

    /// <summary>The document the templates are in.</summary>
    public WordprocessingDocumentWrapper Parent
    {
        get { return _document; }
    }

    private List<AbstractNum> GetAbstractNums()
    {
        Numbering numbering = _document.WordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
        return numbering == null ? new List<AbstractNum>() : numbering.Elements<AbstractNum>().ToList();
    }

    /// <summary>The number of templates.</summary>
    public int Count
    {
        get { return GetAbstractNums().Count; }
    }

    /// <summary>
    /// Adds a template, as <c>ListTemplates.Add</c> in Word: a single level, numbered 1., 2., 3. with the number 0.5 inch in and the text
    /// 0.25 inch further, or, with <paramref name="outlineNumbered"/>, the nine levels of an outline (1), a), i), (1) and so on). The levels can
    /// then be changed with <see cref="ListTemplateWrapper.ListLevels"/>. No list uses the template until it is applied.
    /// </summary>
    /// <param name="outlineNumbered">True for a template with nine levels.</param>
    /// <param name="name">The name of the template; none when null.</param>
    public ListTemplateWrapper Add(bool outlineNumbered = false, string name = null)
    {
        Numbering numbering = ListDefaults.GetOrCreateNumbering(_document.WordprocessingDocument.MainDocumentPart);
        AbstractNum template = new AbstractNum() { AbstractNumberId = numbering.Elements<AbstractNum>().Select(a => a.AbstractNumberId?.Value ?? -1).DefaultIfEmpty(-1).Max() + 1 };
        template.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
        template.Nsid = new Nsid() { Val = ListDefaults.NewNsid() };
        template.MultiLevelType = new MultiLevelType() { Val = outlineNumbered ? MultiLevelValues.Multilevel : MultiLevelValues.SingleLevel };
        template.TemplateCode = new TemplateCode() { Val = ListDefaults.NewNsid() };
        if (!string.IsNullOrEmpty(name))
        {
            template.AbstractNumDefinitionName = new AbstractNumDefinitionName() { Val = name };
        }

        if (outlineNumbered)
        {
            foreach (Level level in OutlineLevels())
            {
                template.AppendChild(level);
            }
        }
        else
        {
            template.AppendChild(NewLevel(0, NumberFormatValues.Decimal, "%1.", 720));
        }

        NumberingInstance firstNum = numbering.Elements<NumberingInstance>().FirstOrDefault();
        if (firstNum != null)
        {
            firstNum.InsertBeforeSelf(template);
        }
        else
        {
            numbering.AppendChild(template);
        }
        return new ListTemplateWrapper(template);
    }

    //The nine levels of a new outline.
    internal static List<Level> OutlineLevels()
    {
        NumberFormatValues[] formats = { NumberFormatValues.Decimal, NumberFormatValues.LowerLetter, NumberFormatValues.LowerRoman };
        string[] texts = { "%{0})", "%{0})", "%{0})", "(%{0})", "(%{0})", "(%{0})", "%{0}.", "%{0}.", "%{0}." };
        List<Level> levels = new List<Level>();
        for (int level = 0; level < 9; level++)
        {
            levels.Add(NewLevel(level, formats[level % 3], texts[level].Replace("{0}", (level + 1).ToString(CultureInfo.InvariantCulture)),
                level == 0 ? 360 : 360 * (level + 1)));
        }
        return levels;
    }

    private static Level NewLevel(int index, NumberFormatValues format, string text, int left)
    {
        return new Level(
            new StartNumberingValue() { Val = 1 },
            new NumberingFormat() { Val = format },
            new LevelText() { Val = text },
            new LevelJustification() { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(new Indentation() { Left = left.ToString(CultureInfo.InvariantCulture), Hanging = "360" }))
        { LevelIndex = index };
    }

    /// <summary>The template at the index (0-based), as <c>ListTemplates.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no template at the index.</exception>
    public ListTemplateWrapper this[int index]
    {
        get
        {
            List<AbstractNum> all = GetAbstractNums();
            if (index < 0 || index >= all.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new ListTemplateWrapper(all[index]);
        }
    }

    /// <inheritdoc/>
    public IEnumerator<ListTemplateWrapper> GetEnumerator()
    {
        return GetAbstractNums().Select(a => new ListTemplateWrapper(a)).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

/// <summary>
/// The levels of a list template, as <c>ListLevels</c> in Word.
/// </summary>
public class ListLevelWrapperCollection : IEnumerable<ListLevelWrapper>
{
    private readonly ListTemplateWrapper _template;

    internal ListLevelWrapperCollection(ListTemplateWrapper template)
    {
        _template = template;
    }

    /// <summary>The template the levels are of.</summary>
    public ListTemplateWrapper Parent
    {
        get { return _template; }
    }

    private List<Level> GetLevels()
    {
        return _template.AbstractNum.Elements<Level>().ToList();
    }

    /// <summary>The number of levels.</summary>
    public int Count
    {
        get { return GetLevels().Count; }
    }

    /// <summary>The level at the index (0-based, so the first is level 1 in Word), as <c>ListLevels.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no level at the index.</exception>
    public ListLevelWrapper this[int index]
    {
        get
        {
            List<Level> levels = GetLevels();
            if (index < 0 || index >= levels.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new ListLevelWrapper(_template, levels[index]);
        }
    }

    /// <inheritdoc/>
    public IEnumerator<ListLevelWrapper> GetEnumerator()
    {
        return GetLevels().Select(l => new ListLevelWrapper(_template, l)).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

/// <summary>
/// One level of a list template, as <c>ListLevel</c> in Word. Lengths are in points.
/// </summary>
public partial class ListLevelWrapper
{
    //What Word answers where a level has no tab stop.
    private const int Undefined = 9999999;

    private readonly ListTemplateWrapper _template;
    private readonly Level _level;

    internal ListLevelWrapper(ListTemplateWrapper template, Level level)
    {
        _template = template;
        _level = level;
    }

    internal Level Level
    {
        get { return _level; }
    }

    /// <summary>The 0-based position of the level in its template (Word counts from 1).</summary>
    public int Index
    {
        get { return _level.LevelIndex?.Value ?? _template.AbstractNum.Elements<Level>().ToList().IndexOf(_level); }
    }

    /// <summary>The template the level is in.</summary>
    public ListTemplateWrapper Parent
    {
        get { return _template; }
    }

    /// <summary>The text of the number, with <c>%1</c> to <c>%9</c> where the numbers of the levels go, as <c>ListLevel.NumberFormat</c> in Word.</summary>
    public string NumberFormat
    {
        get { return _level.LevelText?.Val?.Value ?? string.Empty; }
        set
        {
            _level.LevelText = new LevelText() { Val = value ?? string.Empty };
        }
    }

    /// <summary>How the number is written, as <c>ListLevel.NumberStyle</c> in Word.</summary>
    public ListNumberStyle NumberStyle
    {
        get { return ToStyle(_level.NumberingFormat?.Val?.Value); }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            _level.NumberingFormat = new NumberingFormat() { Val = ToFormat(value) };
        }
    }

    internal static ListNumberStyle ToStyle(NumberFormatValues? format)
    {
        if (format == NumberFormatValues.UpperRoman)
        {
            return ListNumberStyle.UppercaseRoman;
        }
        if (format == NumberFormatValues.LowerRoman)
        {
            return ListNumberStyle.LowercaseRoman;
        }
        if (format == NumberFormatValues.UpperLetter)
        {
            return ListNumberStyle.UppercaseLetter;
        }
        if (format == NumberFormatValues.LowerLetter)
        {
            return ListNumberStyle.LowercaseLetter;
        }
        if (format == NumberFormatValues.Ordinal)
        {
            return ListNumberStyle.Ordinal;
        }
        if (format == NumberFormatValues.CardinalText)
        {
            return ListNumberStyle.CardinalText;
        }
        if (format == NumberFormatValues.OrdinalText)
        {
            return ListNumberStyle.OrdinalText;
        }
        if (format == NumberFormatValues.DecimalZero)
        {
            return ListNumberStyle.ArabicLZ;
        }
        if (format == NumberFormatValues.Bullet)
        {
            return ListNumberStyle.Bullet;
        }
        return format == NumberFormatValues.None ? ListNumberStyle.None : ListNumberStyle.Arabic;
    }

    private static NumberFormatValues ToFormat(ListNumberStyle style)
    {
        return style == ListNumberStyle.UppercaseRoman ? NumberFormatValues.UpperRoman
            : style == ListNumberStyle.LowercaseRoman ? NumberFormatValues.LowerRoman
            : style == ListNumberStyle.UppercaseLetter ? NumberFormatValues.UpperLetter
            : style == ListNumberStyle.LowercaseLetter ? NumberFormatValues.LowerLetter
            : style == ListNumberStyle.Ordinal ? NumberFormatValues.Ordinal
            : style == ListNumberStyle.CardinalText ? NumberFormatValues.CardinalText
            : style == ListNumberStyle.OrdinalText ? NumberFormatValues.OrdinalText
            : style == ListNumberStyle.ArabicLZ ? NumberFormatValues.DecimalZero
            : style == ListNumberStyle.Bullet ? NumberFormatValues.Bullet
            : style == ListNumberStyle.None ? NumberFormatValues.None : NumberFormatValues.Decimal;
    }

    /// <summary>The number the level starts at, as <c>ListLevel.StartAt</c> in Word (1 by default).</summary>
    public int StartAt
    {
        get { return _level.StartNumberingValue?.Val?.Value ?? 1; }
        set
        {
            Guard.ThrowIfNegative(value);
            _level.StartNumberingValue = new StartNumberingValue() { Val = value };
        }
    }

    private int GetTwips(Func<Indentation, string> read)
    {
        Indentation indentation = _level.PreviousParagraphProperties?.Indentation;
        return indentation != null && int.TryParse(read(indentation), NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) ? twips : 0;
    }

    //The space the number is in: the left indent and the hanging (or first line) indent.
    private int Left
    {
        get { return GetTwips(i => i.Left?.Value ?? i.Start?.Value); }
    }

    private int Hanging
    {
        get { return GetTwips(i => i.Hanging?.Value) - GetTwips(i => i.FirstLine?.Value); }
    }

    /// <summary>Where the number starts, from the margin, in points, as <c>ListLevel.NumberPosition</c> in Word.</summary>
    public double NumberPosition
    {
        get { return (Left - Hanging) / 20.0; }
        set
        {
            SetIndents(value, TextPosition);
        }
    }

    /// <summary>Where the text of the item starts (the left indent), in points, as <c>ListLevel.TextPosition</c> in Word.</summary>
    public double TextPosition
    {
        get { return Left / 20.0; }
        set
        {
            SetIndents(NumberPosition, value);
        }
    }

    private void SetIndents(double number, double text)
    {
        if (double.IsNaN(number) || double.IsNaN(text) || text < 0 || number < -1584 || number > 1584 || text > 1584)
        {
            throw new ArgumentOutOfRangeException(nameof(number), "The positions must be between 0 and 1584 points.");
        }
        PreviousParagraphProperties properties = _level.PreviousParagraphProperties ?? (_level.PreviousParagraphProperties = new PreviousParagraphProperties());
        int left = (int)Math.Round(text * 20, MidpointRounding.AwayFromZero);
        int hanging = left - (int)Math.Round(number * 20, MidpointRounding.AwayFromZero);
        Indentation indentation = properties.Indentation ?? (properties.Indentation = new Indentation());
        indentation.Left = left.ToString(CultureInfo.InvariantCulture);
        indentation.Start = null;
        indentation.FirstLine = null;
        indentation.Hanging = hanging > 0 ? hanging.ToString(CultureInfo.InvariantCulture) : null;
        if (hanging < 0)
        {
            indentation.FirstLine = (-hanging).ToString(CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// The position of the tab stop after the number in points, or 9999999 if the level has none, as <c>ListLevel.TabPosition</c> in Word.
    /// Setting 9999999 takes the tab stop away.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The position is negative or beyond 1584 points (other than 9999999).</exception>
    public double TabPosition
    {
        get
        {
            TabStop tab = _level.PreviousParagraphProperties?.Tabs?.Elements<TabStop>().FirstOrDefault(t => t.Val?.Value == TabStopValues.Number || t.Val?.Value == TabStopValues.Left);
            return tab?.Position?.Value is int twips ? twips / 20.0 : Undefined;
        }
        set
        {
            if (double.IsNaN(value) || (value != Undefined && (value < 0 || value > 1584)))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            PreviousParagraphProperties properties = _level.PreviousParagraphProperties ?? (_level.PreviousParagraphProperties = new PreviousParagraphProperties());
            properties.Tabs = null;
            if (value != Undefined)
            {
                properties.Tabs = new Tabs(new TabStop() { Val = TabStopValues.Number, Position = (int)Math.Round(value * 20, MidpointRounding.AwayFromZero) });
            }
            if (!properties.HasChildren)
            {
                properties.Remove();
            }
        }
    }

    /// <summary>The font of the number, as <c>ListLevel.Font</c> in Word. A bullet or number takes the font of the text unless this sets it.</summary>
    public FontWrapper Font
    {
        get
        {
            return FontWrapper.ForLevel(create =>
            {
                NumberingSymbolRunProperties properties = _level.NumberingSymbolRunProperties;
                if (properties == null && create)
                {
                    properties = new NumberingSymbolRunProperties();
                    _level.NumberingSymbolRunProperties = properties;
                }
                return properties;
            });
        }
    }

    /// <summary>What follows the number, as <c>ListLevel.TrailingCharacter</c> in Word.</summary>
    public TrailingCharacter TrailingCharacter
    {
        get
        {
            LevelSuffixValues? suffix = _level.LevelSuffix?.Val?.Value;
            return suffix == LevelSuffixValues.Space ? TrailingCharacter.Space : suffix == LevelSuffixValues.Nothing ? TrailingCharacter.None : TrailingCharacter.Tab;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            _level.LevelSuffix = value == TrailingCharacter.Tab ? null : new LevelSuffix() { Val = value == TrailingCharacter.Space ? LevelSuffixValues.Space : LevelSuffixValues.Nothing };
        }
    }

    /// <summary>How the number is lined up in the space for it, as <c>ListLevel.Alignment</c> in Word.</summary>
    public ListLevelAlignment Alignment
    {
        get
        {
            LevelJustificationValues? justification = _level.LevelJustification?.Val?.Value;
            return justification == LevelJustificationValues.Center ? ListLevelAlignment.Center : justification == LevelJustificationValues.Right ? ListLevelAlignment.Right : ListLevelAlignment.Left;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            _level.LevelJustification = new LevelJustification()
            {
                Val = value == ListLevelAlignment.Center ? LevelJustificationValues.Center : value == ListLevelAlignment.Right ? LevelJustificationValues.Right : LevelJustificationValues.Left
            };
        }
    }

    /// <summary>
    /// The number of the level after which this one starts again, as <c>ListLevel.ResetOnHigher</c> in Word: by default the level before it
    /// (0 for the first level); 0 for a level that never starts again.
    /// </summary>
    public int ResetOnHigher
    {
        get { return _level.LevelRestart?.Val?.Value ?? Index; }
        set
        {
            if (value < 0 || value > Index)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The level must be 0 or a level before this one.");
            }
            _level.LevelRestart = value == Index ? null : new LevelRestart() { Val = value };
        }
    }

    /// <summary>
    /// The name of the paragraph style that makes a paragraph part of this level, as <c>ListLevel.LinkedStyle</c> in Word; empty if none.
    /// Setting it links the style, which the style then points back at (and a list is made for the template if it has none). An empty
    /// string takes the link away.
    /// </summary>
    /// <exception cref="ArgumentException">There is no paragraph style of that name.</exception>
    public string LinkedStyle
    {
        set { SetLinkedStyle(value); }
        get
        {
            string id = _level.ParagraphStyleIdInLevel?.Val?.Value;
            if (string.IsNullOrEmpty(id))
            {
                return string.Empty;
            }
            Style style = StyleHelper.GetStyles(_template.AbstractNum)?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == id);
            return style == null ? id : StyleHelper.IsCustom(style) ? style.StyleName?.Val?.Value ?? id : BuiltInStyles.GetDisplayName(style.StyleName?.Val?.Value) ?? id;
        }
    }
}
