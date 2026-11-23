using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A field of the document, as <c>Field</c> in Word. Both kinds of field are supported: complex fields (a begin, separate and end
/// character with the field code between them) and simple fields (<c>fldSimple</c>).
/// </summary>
/// <remarks>
/// <see cref="Result"/> is the text stored in the document, which Word shows until the field is updated. <see cref="Update"/> works out the
/// result of the fields that do not depend on the layout of the pages, such as DATE, SEQ, REF and DOCPROPERTY; Word updates the others (PAGE,
/// NUMPAGES, TOC and so on) when the document is opened.
/// </remarks>
public partial class FieldWrapper : System.IEquatable<FieldWrapper>
{
    //Not read-only: a simple field given a result of several paragraphs becomes a complex field (ToComplexField).
    private SimpleField _simpleField;
    private FieldChar _begin;

    public FieldWrapper(SimpleField simpleField)
    {
        Guard.ThrowIfNull(simpleField);
        _simpleField = simpleField;
    }

    /// <param name="begin">The begin character of a complex field.</param>
    public FieldWrapper(FieldChar begin)
    {
        Guard.ThrowIfNull(begin);
        if (!(begin.Parent is Run) || begin.FieldCharType == null || !begin.FieldCharType.HasValue || begin.FieldCharType.Value != FieldCharValues.Begin)
        {
            throw new ArgumentException("The element is not the begin character of a field.", nameof(begin));
        }
        _begin = begin;
    }

    #region Parsing

    private sealed class Parts
    {
        public FieldChar Separate;
        public FieldChar End;
        public readonly List<FieldCode> Codes = new List<FieldCode>();
        public readonly List<Text> ResultText = new List<Text>();
        public readonly List<Run> ResultRuns = new List<Run>();
        public readonly List<Run> AllRuns = new List<Run>();
    }

    //Walks the runs from the begin character to the matching end character, keeping track of nested fields.
    private Parts Parse()
    {
        Parts parts = new Parts();
        Run beginRun = (Run)_begin.Parent;
        int depth = 0;
        bool started = false;
        foreach (Run run in _begin.GetRoot().Descendants<Run>())
        {
            if (!started)
            {
                if (run != beginRun)
                {
                    continue;
                }
                started = true;
            }

            parts.AllRuns.Add(run);
            bool afterSeparate = parts.Separate != null;
            foreach (OpenXmlElement child in run.ChildElements)
            {
                FieldChar fieldChar = child as FieldChar;
                if (fieldChar != null && fieldChar.FieldCharType != null && fieldChar.FieldCharType.HasValue)
                {
                    FieldCharValues type = fieldChar.FieldCharType.Value;
                    if (type == FieldCharValues.Begin)
                    {
                        depth++;
                    }
                    else if (type == FieldCharValues.Separate)
                    {
                        if (depth == 1 && parts.Separate == null)
                        {
                            parts.Separate = fieldChar;
                        }
                    }
                    else if (type == FieldCharValues.End)
                    {
                        depth--;
                        if (depth == 0)
                        {
                            parts.End = fieldChar;
                            return parts;
                        }
                    }
                }
                else if (child is FieldCode && depth == 1 && parts.Separate == null)
                {
                    parts.Codes.Add((FieldCode)child);
                }
                else if (child is Text && parts.Separate != null)
                {
                    parts.ResultText.Add((Text)child);
                }
            }

            if (afterSeparate)
            {
                parts.ResultRuns.Add(run);
            }
        }
        return parts;
    }

    #endregion

    /// <summary>
    /// The characters of the code (<paramref name="code"/>) or of the result of the field as a range, which is empty when there are none. The range of
    /// the code of a simple field (<c>w:fldSimple</c>), which keeps its code in an attribute, is empty at the start of the field.
    /// </summary>
    internal RangeWrapper GetPartRange(bool code)
    {
        OpenXmlElement anchor = _simpleField != null ? (OpenXmlElement)_simpleField : _begin;
        OpenXmlElement story = StoryMap.FindRoot(anchor);
        List<Run> runs;
        OpenXmlElement emptyAfter;
        if (_simpleField != null)
        {
            runs = code ? new List<Run>() : _simpleField.Elements<Run>().ToList();
            emptyAfter = _simpleField;
        }
        else
        {
            Parts parts = Parse();
            if (!code && parts.Separate != null && parts.End != null)
            {
                //The result is everything between the separator and the end, paragraph marks too, as in Word.
                StoryMap map = StoryMap.Build(story);
                if (map.TryGetAtom(parts.Separate, out Atom separate) && map.TryGetAtom(parts.End, out Atom end) && separate.End <= end.Start)
                {
                    return RangeWrapper.FromMap(map, separate.End, end.Start);
                }
            }
            runs = code ? parts.Codes.Select(c => (Run)c.Parent).Distinct().ToList() : parts.ResultRuns;
            emptyAfter = code ? _begin.Parent : parts.Separate?.Parent ?? _begin.Parent;
        }

        RangeWrapper range = new RangeWrapper(story);
        if (runs.Count > 0)
        {
            range.SetRange(new RangeWrapper(runs[0]).Start, new RangeWrapper(runs[runs.Count - 1]).End);
        }
        else
        {
            //A simple field has no characters of its own to be after, so its empty code is at its start.
            int position = _simpleField != null ? new RangeWrapper(_simpleField).Start : new RangeWrapper(emptyAfter).End;
            range.SetRange(position, position);
        }
        return range;
    }

    /// <summary>
    /// The field code without its surrounding spaces, such as <c>PAGE \* MERGEFORMAT</c>. Setting it replaces the code.
    /// </summary>
    internal string CodeText
    {
        get
        {
            if (_simpleField != null)
            {
                return (_simpleField.Instruction?.Value ?? string.Empty).Trim();
            }
            return string.Concat(Parse().Codes.Select(c => c.Text)).Trim();
        }
        set
        {
            Guard.ThrowIfNull(value);
            string code = " " + value.Trim() + " ";
            if (_simpleField != null)
            {
                _simpleField.Instruction = code;
                return;
            }

            Parts parts = Parse();
            if (parts.Codes.Count == 0)
            {
                ((Run)_begin.Parent).InsertAfterSelf(new Run(new FieldCode(code) { Space = SpaceProcessingModeValues.Preserve }));
                return;
            }
            parts.Codes[0].Text = code;
            parts.Codes[0].Space = SpaceProcessingModeValues.Preserve;
            foreach (FieldCode other in parts.Codes.Skip(1))
            {
                OpenXmlElement run = other.Parent;
                other.Remove();
                if (run is Run && !run.Elements().Any(e => !(e is RunProperties)))
                {
                    run.Remove();
                }
            }
        }
    }

    /// <summary>
    /// The kind of field, worked out from the first word of the code.
    /// </summary>
    public FieldType Type
    {
        get
        {
            string code = CodeText;
            int space = code.IndexOf(' ', StringComparison.Ordinal);
            return FieldCodes.FromName(space < 0 ? code : code.Substring(0, space));
        }
    }

    /// <summary>
    /// The text stored as the result of the field, as <c>Field.Result.Text</c> in Word, with a paragraph mark (\r) between the paragraphs of a
    /// result of several. Setting it replaces the result, keeping the formatting of its first run; a paragraph mark in the text starts a new
    /// paragraph in the formatting of the field's own paragraph, and text without one joins the paragraphs of a result of several again, as in
    /// Word.
    /// </summary>
    internal string ResultText
    {
        get
        {
            if (_simpleField != null)
            {
                return string.Concat(_simpleField.Descendants<Text>().Select(t => t.Text));
            }
            Parts parts = Parse();
            if (parts.Separate != null && parts.End != null && !InOneParagraph(parts))
            {
                return GetPartRange(false).Text;
            }
            return string.Concat(parts.ResultText.Select(t => t.Text));
        }
        set
        {
            string text = value ?? string.Empty;
            if (_simpleField != null && text.IndexOf('\r') < 0)
            {
                SetSimpleResult(text);
                return;
            }
            if (_simpleField != null)
            {
                ToComplexField();
            }
            SetComplexResult(text);
        }
    }

    //Whether the separator, the result and the end of the field are all in one paragraph.
    private static bool InOneParagraph(Parts parts)
    {
        OpenXmlElement paragraph = parts.End?.Parent?.Parent;
        return (parts.Separate == null || parts.Separate.Parent.Parent == paragraph) && parts.ResultRuns.All(r => r.Parent == paragraph);
    }

    //Writes a simple field (w:fldSimple) as a field of begin, code, separator, result and end, as Word does once it has the field.
    private void ToComplexField()
    {
        SimpleField simple = _simpleField;
        //The begin, code, separator and end are in the formatting of the result, but not its "do not check spelling", which is the result's own.
        RunProperties properties = simple.Elements<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        properties?.RemoveAllChildren<NoProof>();
        Run Mark(OpenXmlElement content)
        {
            Run run = new Run(content);
            if (properties != null && properties.HasChildren)
            {
                run.PrependChild(properties.CloneNode(true));
            }
            return run;
        }

        FieldChar begin = new FieldChar() { FieldCharType = FieldCharValues.Begin };
        if (simple.FieldLock != null)
        {
            begin.FieldLock = simple.FieldLock.Value;
        }
        if (simple.Dirty != null)
        {
            begin.Dirty = simple.Dirty.Value;
        }
        List<OpenXmlElement> runs = new List<OpenXmlElement>
        {
            Mark(begin),
            Mark(new FieldCode(simple.Instruction?.Value ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }),
            Mark(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
        };
        foreach (OpenXmlElement child in simple.ChildElements.Where(c => !(c is FieldData)).ToList())
        {
            child.Remove();
            runs.Add(child);
        }
        runs.Add(Mark(new FieldChar() { FieldCharType = FieldCharValues.End }));
        foreach (OpenXmlElement run in runs)
        {
            simple.InsertBeforeSelf(run);
        }
        simple.Remove();
        _simpleField = null;
        _begin = begin;
    }

    //True while the result is set through Result.Text: Word does not keep the "do not check spelling" of a result it worked out for text put in
    //its place.
    private bool _settingText;

    internal void SetResultText(string text)
    {
        _settingText = true;
        try
        {
            ResultText = text;
        }
        finally
        {
            _settingText = false;
        }
    }

    private RunProperties WithoutProofing(RunProperties properties)
    {
        if (_settingText && properties != null)
        {
            properties.RemoveAllChildren<NoProof>();
            if (!properties.HasChildren)
            {
                return null;
            }
        }
        return properties;
    }

    private void SetSimpleResult(string text)
    {
        Run first = _simpleField.Elements<Run>().FirstOrDefault();
        RunProperties properties = WithoutProofing(first?.RunProperties?.CloneNode(true) as RunProperties);
        _simpleField.RemoveAllChildren<Run>();
        Run run = new Run(OpenXmlExtensions.NewText(text));
        if (properties != null)
        {
            run.PrependChild(properties);
        }
        _simpleField.AppendChild(run);
    }

    //Takes away a result of several paragraphs, so that the field's end, and whatever follows it in its paragraph, come back to the paragraph of
    //the separator, which keeps its formatting (and the section break of the end's paragraph, if it ended one), as deleting the result does in
    //Word.
    private void JoinResult(Parts parts)
    {
        Run separateRun = (Run)parts.Separate.Parent;
        Run endRun = (Run)parts.End.Parent;
        Paragraph first = separateRun.Ancestors<Paragraph>().First();
        Paragraph last = endRun.Ancestors<Paragraph>().First();
        if (separateRun.Parent != first || endRun.Parent != last)
        {
            throw new NotSupportedException("The result of a field can only be replaced when its separator and end are directly in paragraphs.");
        }

        List<OpenXmlElement> going = separateRun.ElementsAfter().ToList();
        for (OpenXmlElement between = first.NextSibling(); between != null && between != last; between = between.NextSibling())
        {
            going.Add(between);
        }
        going.AddRange(endRun.ElementsBefore().Where(e => !(e is ParagraphProperties)));
        if (going.Any(e => e is Table || e is SdtBlock || (!(e is Paragraph) && !(e.Parent is Paragraph))))
        {
            throw new NotSupportedException("The result of a field can only be replaced when it is made of paragraphs.");
        }

        HashSet<OpenXmlElement> removed = new HashSet<OpenXmlElement>(going.SelectMany(e => e.Descendants().Prepend(e)));
        removed.Add(last);
        StoryMap map = StoryMap.Build(StoryMap.FindRoot(first));
        Anchor after = map.TryGetAtom(parts.Separate, out Atom separate) ? Anchor.At(map, separate.End) : null;
        if (after != null)
        {
            SpanRegistry.Removing(first, removed, after);
        }
        foreach (OpenXmlElement element in going)
        {
            element.Remove();
        }

        SectionProperties section = last.ParagraphProperties?.GetFirstChild<SectionProperties>();
        foreach (OpenXmlElement element in last.ChildElements.Where(e => !(e is ParagraphProperties)).ToList())
        {
            element.Remove();
            first.AppendChild(element);
        }
        if (section != null)
        {
            section.Remove();
            first.GetOrAddParagraphProperties().AppendChild(section);
        }
        SpanRegistry.ParagraphMerged(last, first);
        last.Remove();
    }

    private void SetComplexResult(string text)
    {
        Parts parts = Parse();
        if (parts.End == null)
        {
            throw new InvalidOperationException("The field has no end character.");
        }
        Run endRun = (Run)parts.End.Parent;
        Run separateRun = parts.Separate == null ? null : (Run)parts.Separate.Parent;
        if (!InOneParagraph(parts))
        {
            //A result of several paragraphs is first joined into one, as deleting it does in Word, keeping the formatting of its first run.
            RunProperties first = WithoutProofing(parts.ResultRuns.FirstOrDefault(r => r.Elements<Text>().Any())?.RunProperties?.CloneNode(true) as RunProperties);
            JoinResult(parts);
            SetComplexResult(string.Empty);
            Run joined = Parse().ResultRuns.FirstOrDefault();
            if (first != null && joined != null)
            {
                joined.RunProperties?.Remove();
                joined.PrependChild(first);
            }
            SetComplexResult(text);
            return;
        }
        int mark = text.IndexOf('\r');
        if (mark >= 0)
        {
            //The first paragraph of the text goes where the result is; the rest goes in before the end, splitting the paragraph.
            SetComplexResult(text.Substring(0, mark));
            RangeWrapper result = GetPartRange(false);
            result.Collapse(CollapseDirection.End);
            result.InsertAfter(text.Substring(mark));
            return;
        }

        //A field with no result yet has its result in the formatting of the field.
        RunProperties properties = WithoutProofing((parts.ResultRuns.FirstOrDefault()?.RunProperties ?? ((Run)_begin.Parent).RunProperties)?.CloneNode(true) as RunProperties);
        foreach (Run old in parts.ResultRuns)
        {
            old.Remove();
        }

        if (separateRun == null)
        {
            //A field without a result gets the separator that introduces one.
            separateRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
            endRun.InsertBeforeSelf(separateRun);
        }

        Run resultRun = new Run(OpenXmlExtensions.NewText(text));
        if (properties != null)
        {
            resultRun.PrependChild(properties);
        }
        separateRun.InsertAfterSelf(resultRun);
    }

    //Word does not check the spelling of the result of a field it has just worked out.
    internal void MarkResultNotProofed()
    {
        Run result = _simpleField != null ? _simpleField.Elements<Run>().FirstOrDefault() : Parse().ResultRuns.FirstOrDefault();
        result?.GetOrAddRunProperties().SetRunProperty(new NoProof());
    }

    //Writes a new field as one element, as Word does when it has a result and no formatting of its own. Returns the field as it now is.
    internal FieldWrapper ToSimpleFieldIfPlain()
    {
        Parts parts = Parse();
        if (_begin == null || parts.AllRuns.Any(r => r.RunProperties != null && r.RunProperties.HasChildren && !parts.ResultRuns.Contains(r)))
        {
            return this;
        }

        //Word writes the formulas of table cells in pieces.
        string codeText = string.Concat(parts.Codes.Select(c => c.Text)).TrimStart();
        if (codeText.Length > 0 && codeText[0] == '=')
        {
            return this;
        }

        Run beginRun = (Run)_begin.Parent;
        SimpleField simple = new SimpleField() { Instruction = string.Concat(parts.Codes.Select(c => c.Text)) };
        foreach (Run result in parts.ResultRuns)
        {
            result.Remove();
            simple.AppendChild(result);
        }
        beginRun.InsertBeforeSelf(simple);
        foreach (Run run in parts.AllRuns.Where(r => r.Parent != null && r.Parent != simple).ToList())
        {
            run.Remove();
        }
        return new FieldWrapper(simple);
    }

    /// <summary>
    /// The paragraph that contains the start of the field.
    /// </summary>
    public RangeWrapper Parent
    {
        get { return new RangeWrapper(Key.Ancestors<Paragraph>().First()); }
    }

    /// <summary>
    /// A range from the first to the last element of the field.
    /// </summary>
    public RangeWrapper Range
    {
        get
        {
            if (_simpleField != null)
            {
                return new RangeWrapper(_simpleField);
            }
            Parts parts = Parse();
            return new RangeWrapper(_begin.Parent, parts.End?.Parent ?? parts.AllRuns.Last());
        }
    }

    /// <summary>
    /// Removes the field, including its result.
    /// </summary>
    public void Delete()
    {
        if (_simpleField != null)
        {
            _simpleField.Remove();
            return;
        }

        foreach (Run run in Parse().AllRuns)
        {
            run.Remove();
        }
    }

    /// <summary>Selects the field, as <c>Field.Select</c> in Word.</summary>
    public void Select()
    {
        Range.Select();
    }

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    internal OpenXmlElement Key
    {
        get { return _simpleField ?? (OpenXmlElement)_begin; }
    }

    public override bool Equals(object obj)
    {
        return Equals(obj as FieldWrapper);
    }

    public bool Equals(FieldWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(Key, other.Key);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(Key);
    }

    public static bool operator ==(FieldWrapper left, FieldWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(FieldWrapper left, FieldWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
