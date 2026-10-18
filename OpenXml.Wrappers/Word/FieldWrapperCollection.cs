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
    public FieldWrapper Add(RangeWrapper range, FieldType type, string text = null)
    {
        string code = FieldWrapper.TypeToCodeName(type);
        return Add(range, String.IsNullOrWhiteSpace(text) ? code : code + " " + text.Trim());
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
        return new FieldWrapper(begin);
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
