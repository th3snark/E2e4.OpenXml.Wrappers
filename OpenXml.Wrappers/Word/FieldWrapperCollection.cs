using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The fields in a range, as <c>Fields</c> in Word. Fields nested in the result of another field are included, in document order.
/// </summary>
public class FieldWrapperCollection : IEnumerable<FieldWrapper>
{
    private readonly RangeWrapper _range;

    public FieldWrapperCollection(RangeWrapper range)
    {
        Guard.ThrowIfNull(range);
        _range = range;
    }

    /// <summary>
    /// The range the collection was taken from.
    /// </summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    public FieldWrapper this[int index]
    {
        get
        {
            IList<FieldWrapper> fields = GetElementList();
            if (index < 0 || index >= fields.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return fields[index];
        }
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    private IList<FieldWrapper> GetElementList()
    {
        List<OpenXmlElement> starts = new List<OpenXmlElement>();
        starts.AddRange(_range.StartElement.GetElements<Run>(_range.EndElement)
            .SelectMany(r => r.Elements<FieldChar>())
            .Where(f => f.FieldCharType != null && f.FieldCharType.HasValue && f.FieldCharType.Value == FieldCharValues.Begin));
        starts.AddRange(_range.StartElement.GetElements<SimpleField>(_range.EndElement));

        //Complex fields and simple fields are found separately, so put them back in document order.
        starts.Sort((a, b) => ReferenceEquals(a, b) ? 0 : a.IsBefore(b) ? -1 : 1);
        return starts.Select(e => e is SimpleField ? new FieldWrapper((SimpleField)e) : new FieldWrapper((FieldChar)e)).ToList();
    }

    /// <summary>
    /// Adds a field of a known type at the end of the range, as <c>Fields.Add</c> in Word, and returns it. It is marked so that Word updates
    /// it when the document is opened.
    /// </summary>
    /// <param name="range">Where to add the field. The rules are those of <see cref="RangeWrapper.InsertAfter"/>.</param>
    /// <param name="type">The type of field.</param>
    /// <param name="text">The rest of the field code (the arguments and switches), or null.</param>
    /// <param name="preserveFormatting">True (as in Word) to add <c>* MERGEFORMAT</c> so that formatting given to the result is kept when the field is updated.</param>
    public FieldWrapper Add(RangeWrapper range, FieldType type = FieldType.Empty, string text = null, bool preserveFormatting = true)
    {
        string code;
        if (type == FieldType.Empty)
        {
            code = text;
        }
        else
        {
            string name = FieldCodes.ToName(type);
            code = String.IsNullOrWhiteSpace(text) ? name : name + " " + text.Trim();
        }
        if (preserveFormatting && !String.IsNullOrWhiteSpace(code) && !code.Contains("MERGEFORMAT", StringComparison.OrdinalIgnoreCase))
        {
            code = code.Trim() + " \\* MERGEFORMAT";
        }
        return Add(range, code);
    }

    /// <summary>
    /// Adds a field with the given code, for example <c>PAGE \* MERGEFORMAT</c>, at the end of the range.
    /// </summary>
    public FieldWrapper Add(RangeWrapper range, string code)
    {
        Guard.ThrowIfNull(range);
        if (String.IsNullOrWhiteSpace(code))
        {
            throw new ArgumentException("A field needs a code.", nameof(code));
        }

        FieldChar begin = new FieldChar() { FieldCharType = FieldCharValues.Begin, Dirty = true };
        range.InsertFieldRuns(
            _ => new List<Run>
            {
                new Run(begin),
                new Run(new FieldCode(" " + code.Trim() + " ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(String.Empty) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.End })
            });
        FieldWrapper field = new FieldWrapper(begin);
        //As in Word a new field shows its result straight away when that can be worked out; the others are left for Word to update.
        field.Update();
        return field;
    }

    /// <summary>
    /// Updates the fields, as <c>Fields.Update</c> in Word, and returns 0 if all of them were updated, otherwise the number that could not be
    /// (Word returns the position of the first). See <see cref="FieldWrapper.Update"/>.
    /// </summary>
    public int Update()
    {
        return GetElementList().Count(f => !f.Update());
    }

    /// <summary>Locks every field so that it is not updated (<c>Fields.Lock</c> in Word).</summary>
    public void Lock()
    {
        foreach (FieldWrapper field in GetElementList())
        {
            field.Locked = true;
        }
    }

    /// <summary>Unlocks every field (<c>Fields.Unlock</c> in Word).</summary>
    public void Unlock()
    {
        foreach (FieldWrapper field in GetElementList())
        {
            field.Locked = false;
        }
    }

    /// <summary>True if every field is locked, false if none is, and null if some are.</summary>
    public bool? Locked
    {
        get
        {
            List<bool> states = GetElementList().Select(f => f.Locked).Distinct().ToList();
            return states.Count == 1 ? states[0] : states.Count == 0 ? false : (bool?)null;
        }
    }

    #region IEnumerable Members

    public IEnumerator<FieldWrapper> GetEnumerator()
    {
        foreach (FieldWrapper field in GetElementList())
        {
            yield return field;
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
