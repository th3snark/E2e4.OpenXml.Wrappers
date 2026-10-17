using System;
using System.Globalization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The type of the value of a custom document property, as <c>MsoDocProperties</c> in Office.
/// </summary>
public enum DocumentPropertyType
{
    Boolean = 0,
    Date = 1,
    FloatingPoint = 2,
    Number = 3,
    Text = 4
}

/// <summary>
/// A custom property of the document, as one member of <c>Document.CustomDocumentProperties</c> in Word. The value is a
/// <see cref="bool"/>, <see cref="DateTime"/> (UTC), <see cref="double"/>, <see cref="int"/> or <see cref="string"/> by <see cref="Type"/>.
/// </summary>
public class CustomDocumentProperty : System.IEquatable<CustomDocumentProperty>
{
    private readonly DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty _property;

    internal CustomDocumentProperty(DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty property)
    {
        _property = property;
    }

    public DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty Property
    {
        get { return _property; }
    }

    public string Name
    {
        get { return _property.Name?.Value; }
    }

    public DocumentPropertyType Type
    {
        get
        {
            if (_property.VTBool != null)
            {
                return DocumentPropertyType.Boolean;
            }
            if (_property.VTFileTime != null)
            {
                return DocumentPropertyType.Date;
            }
            if (_property.VTDouble != null)
            {
                return DocumentPropertyType.FloatingPoint;
            }
            return _property.VTInt32 != null ? DocumentPropertyType.Number : DocumentPropertyType.Text;
        }
    }

    /// <summary>
    /// The value. Setting a value of another kind changes the <see cref="Type"/> of the property.
    /// </summary>
    /// <exception cref="ArgumentException">The value is not a bool, DateTime, double, int or string (or null, which is an empty string).</exception>
    public object Value
    {
        get
        {
            switch (Type)
            {
                case DocumentPropertyType.Boolean:
                    return String.Equals(_property.VTBool.Text, "true", StringComparison.OrdinalIgnoreCase) || _property.VTBool.Text == "1";
                case DocumentPropertyType.Date:
                    return DateTime.Parse(_property.VTFileTime.Text, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal);
                case DocumentPropertyType.FloatingPoint:
                    return Double.Parse(_property.VTDouble.Text, NumberStyles.Float, CultureInfo.InvariantCulture);
                case DocumentPropertyType.Number:
                    return Int32.Parse(_property.VTInt32.Text, NumberStyles.Integer, CultureInfo.InvariantCulture);
                default:
                    return _property.VTLPWSTR?.Text ?? String.Empty;
            }
        }
        set
        {
            _property.RemoveAllChildren();
            _property.AppendChild(CreateValue(value));
        }
    }

    internal static OpenXmlElement CreateValue(object value)
    {
        if (value == null)
        {
            return new VTLPWSTR(String.Empty);
        }
        if (value is string text)
        {
            return new VTLPWSTR(text);
        }
        if (value is bool flag)
        {
            return new VTBool(flag ? "true" : "false");
        }
        if (value is int number)
        {
            return new VTInt32(number.ToString(CultureInfo.InvariantCulture));
        }
        if (value is double || value is float)
        {
            return new VTDouble(Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture));
        }
        if (value is DateTime date)
        {
            return new VTFileTime(date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
        }
        throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "A value of the type {0} cannot be stored in a document property.", value.GetType().Name), nameof(value));
    }

    /// <summary>
    /// Removes the property from the document.
    /// </summary>
    public void Delete()
    {
        _property.Remove();
    }

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as CustomDocumentProperty);
    }

    public bool Equals(CustomDocumentProperty other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_property, other._property);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_property);
    }

    public static bool operator ==(CustomDocumentProperty left, CustomDocumentProperty right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(CustomDocumentProperty left, CustomDocumentProperty right)
    {
        return !(left == right);
    }

    #endregion
}
