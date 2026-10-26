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
    private readonly SimpleField _simpleField;
    private readonly FieldChar _begin;

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
    /// The text stored as the result of the field, as <c>Field.Result.Text</c> in Word. Setting it replaces the result, keeping the formatting
    /// of its first run.
    /// </summary>
    /// <exception cref="NotSupportedException">The result is not within a single paragraph.</exception>
    internal string ResultText
    {
        get
        {
            if (_simpleField != null)
            {
                return string.Concat(_simpleField.Descendants<Text>().Select(t => t.Text));
            }
            return string.Concat(Parse().ResultText.Select(t => t.Text));
        }
        set
        {
            string text = value ?? string.Empty;
            if (_simpleField != null)
            {
                SetSimpleResult(text);
            }
            else
            {
                SetComplexResult(text);
            }
        }
    }

    private void SetSimpleResult(string text)
    {
        Run first = _simpleField.Elements<Run>().FirstOrDefault();
        RunProperties properties = first?.RunProperties?.CloneNode(true) as RunProperties;
        _simpleField.RemoveAllChildren<Run>();
        Run run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        if (properties != null)
        {
            run.PrependChild(properties);
        }
        _simpleField.AppendChild(run);
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
        if ((separateRun != null && separateRun.Parent != endRun.Parent) || parts.ResultRuns.Any(r => r.Parent != endRun.Parent))
        {
            throw new NotSupportedException("The result of a field can only be set when it is within a single paragraph.");
        }

        RunProperties properties = parts.ResultRuns.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
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

        Run resultRun = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        if (properties != null)
        {
            resultRun.PrependChild(properties);
        }
        separateRun.InsertAfterSelf(resultRun);
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
