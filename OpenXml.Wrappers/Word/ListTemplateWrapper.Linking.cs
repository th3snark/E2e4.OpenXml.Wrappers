using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class ListLevelWrapper
{
    //Links the level to a paragraph style, as Word does: the level names the style, and the style points back at a list that uses the
    //template (one is made if there is none) and at the level.
    private void SetLinkedStyle(string name)
    {
        Guard.ThrowIfNull(name);
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(_template.AbstractNum) ?? throw new NotSupportedException("A level can be linked to a style only in a document.");
        Numbering numbering = (Numbering)_template.AbstractNum.Parent;

        string oldId = _level.ParagraphStyleIdInLevel?.Val?.Value;
        if (!string.IsNullOrEmpty(oldId))
        {
            Style old = main.StyleDefinitionsPart?.Styles?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == oldId);
            old?.StyleParagraphProperties?.NumberingProperties?.Remove();
        }
        _level.ParagraphStyleIdInLevel = null;
        if (name.Length == 0)
        {
            return;
        }

        Style style = BuiltInStyles.GetOrCreate(main, name);
        if (style == null || style.Type?.Value != StyleValues.Paragraph)
        {
            throw new ArgumentException("There is no paragraph style with that name.", nameof(name));
        }
        //A style belongs to one level only: another level that names it lets go of it.
        foreach (Level other in numbering.Descendants<Level>().Where(l => l != _level && l.ParagraphStyleIdInLevel?.Val?.Value == style.StyleId?.Value).ToList())
        {
            other.ParagraphStyleIdInLevel = null;
        }
        _level.ParagraphStyleIdInLevel = new ParagraphStyleIdInLevel() { Val = style.StyleId.Value };

        int abstractId = _template.AbstractNum.AbstractNumberId.Value;
        NumberingInstance num = numbering.Elements<NumberingInstance>().FirstOrDefault(n => n.AbstractNumId?.Val?.Value == abstractId)
            ?? ListDefaults.AddNum(numbering, _template.AbstractNum);
        NumberingProperties properties = new NumberingProperties() { NumberingId = new NumberingId() { Val = num.NumberID.Value } };
        if (Index > 0)
        {
            properties.NumberingLevelReference = new NumberingLevelReference() { Val = Index };
        }
        StyleParagraphProperties paragraph = style.StyleParagraphProperties ?? (style.StyleParagraphProperties = new StyleParagraphProperties());
        paragraph.NumberingProperties = properties;
    }
}

public partial class ListTemplateWrapper
{
    /// <summary>
    /// Converts a template with several levels to one with a single level, or the other way round, as <c>ListTemplate.Convert</c> in Word.
    /// A template with several levels becomes one with the level given as its only level (formatted as that level was); a single-level
    /// template becomes an outline whose other eight levels have the defaults of a new outline.
    /// </summary>
    /// <param name="level">The level to keep (1 to 9) when converting an outline to a single level; 1 is the only valid value the other way.</param>
    /// <exception cref="ArgumentOutOfRangeException">The level is not valid.</exception>
    public void Convert(int level = 1)
    {
        if (OutlineNumbered)
        {
            if (level < 1 || level > 9)
            {
                throw new ArgumentOutOfRangeException(nameof(level));
            }
            Level keep = _abstractNum.Elements<Level>().First(l => (l.LevelIndex?.Value ?? 0) == level - 1);
            foreach (Level other in _abstractNum.Elements<Level>().Where(l => l != keep).ToList())
            {
                other.Remove();
            }
            keep.LevelIndex = 0;
            _abstractNum.MultiLevelType = new MultiLevelType() { Val = MultiLevelValues.SingleLevel };
            return;
        }

        if (level != 1)
        {
            throw new ArgumentOutOfRangeException(nameof(level), "A single-level template can be converted only with level 1.");
        }
        List<Level> outline = ListTemplateWrapperCollection.OutlineLevels();
        Level last = _abstractNum.Elements<Level>().Last();
        foreach (Level added in outline.Skip(1))
        {
            last.InsertAfterSelf(added);
            last = added;
        }
        _abstractNum.MultiLevelType = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
    }
}
