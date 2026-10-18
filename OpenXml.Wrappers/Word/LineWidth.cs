using System;
using System.Globalization;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The widths of line that Word names, with the values of <c>WdLineWidth</c> (eighths of a point).
/// </summary>
public enum BorderLineWidth
{
    Pt025 = 2,
    Pt050 = 4,
    Pt075 = 6,
    Pt100 = 8,
    Pt150 = 12,
    Pt225 = 18,
    Pt300 = 24,
    Pt450 = 36,
    Pt600 = 48
}

/// <summary>
/// The width of a border line in points, from 0.25 to 12 (rounded to eighths of a point), as the <c>LineWidth</c> of Word's <c>Border</c>.
/// </summary>
/// <remarks>
/// A width converts implicitly from a number of points and from a <see cref="BorderLineWidth"/>, so that both
/// <c>border.LineWidth = 1.5</c> and <c>border.LineWidth = BorderLineWidth.Pt150</c> work. The default value, which is what a border that has not been
/// set (or borders that differ) read as, is a width of 0 points.
/// </remarks>
public readonly struct LineWidth : IEquatable<LineWidth>
{
    /// <summary>The narrowest line, in points.</summary>
    public const double MinimumPoints = 0.25;

    /// <summary>The widest line, in points.</summary>
    public const double MaximumPoints = 12;

    /// <summary>The width of a border that has not been set, and what borders that differ read as: 0 points. It cannot be assigned to a border.</summary>
    public static LineWidth None
    {
        get { return default(LineWidth); }
    }

    /// <summary>Makes a width from a number of points.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is not from 0.25 to 12 points.</exception>
    public LineWidth(double points)
    {
        if (Double.IsNaN(points) || points < MinimumPoints || points > MaximumPoints)
        {
            throw new ArgumentOutOfRangeException(nameof(points), String.Format(CultureInfo.InvariantCulture, "The width must be between {0} and {1} points.", MinimumPoints, MaximumPoints));
        }
        Points = Math.Round(points * 8, MidpointRounding.AwayFromZero) / 8;
    }

    /// <summary>The width in points; 0 for a border that has not been set.</summary>
    public double Points { get; }

    /// <summary>The width in eighths of a point, the unit Open XML stores.</summary>
    internal uint EighthPoints
    {
        get { return (uint)Math.Round(Points * 8, MidpointRounding.AwayFromZero); }
    }

    internal static LineWidth FromEighthPoints(uint eighths)
    {
        return eighths == 0 ? default(LineWidth) : new LineWidth(Math.Min(MaximumPoints, Math.Max(MinimumPoints, eighths / 8.0)));
    }

    /// <summary>Converts a number of points to a width.</summary>
    public static implicit operator LineWidth(double points)
    {
        return new LineWidth(points);
    }

    /// <summary>Converts a named width to a width.</summary>
    public static implicit operator LineWidth(BorderLineWidth width)
    {
        return new LineWidth((int)width / 8.0);
    }

    /// <inheritdoc />
    public bool Equals(LineWidth other)
    {
        return Points == other.Points;
    }

    /// <inheritdoc />
    public override bool Equals(object obj)
    {
        return obj is LineWidth other && Equals(other);
    }

    /// <inheritdoc />
    public override int GetHashCode()
    {
        return Points.GetHashCode();
    }

    /// <summary>Whether two widths are the same.</summary>
    public static bool operator ==(LineWidth left, LineWidth right)
    {
        return left.Equals(right);
    }

    /// <summary>Whether two widths are different.</summary>
    public static bool operator !=(LineWidth left, LineWidth right)
    {
        return !left.Equals(right);
    }

    /// <summary>The width in points.</summary>
    public override string ToString()
    {
        return Points.ToString(CultureInfo.InvariantCulture);
    }
}
