using System;
using System.Globalization;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A colour by name, with the values of <c>WdColor</c> in Word (a blue, green and red value packed as <c>0xBBGGRR</c>, the order the
/// VBA function <c>RGB</c> uses).
/// </summary>
public enum NamedColor
{
    Automatic = -16777216,
    Black = 0,
    Blue = 16711680,
    Turquoise = 16776960,
    BrightGreen = 65280,
    Pink = 16711935,
    Red = 255,
    Yellow = 65535,
    White = 16777215,
    DarkBlue = 8388608,
    Teal = 8421376,
    Green = 32768,
    Violet = 8388736,
    DarkRed = 128,
    DarkYellow = 32896,
    Gray50 = 8421504,
    Gray25 = 12632256,
    Aqua = 13421619,
    BlueGray = 10053222,
    Brown = 13209,
    DarkGreen = 13056,
    DarkTeal = 6697728,
    Gold = 52479,
    Indigo = 10040115,
    Lavender = 16751052,
    LightBlue = 16737843,
    LightGreen = 13434828,
    LightOrange = 39423,
    LightTurquoise = 16777164,
    LightYellow = 10092543,
    Lime = 52377,
    OliveGreen = 13107,
    Orange = 26367,
    PaleBlue = 16764057,
    Plum = 6697881,
    Rose = 13408767,
    SeaGreen = 6723891,
    SkyBlue = 16763904,
    Tan = 10079487
}

/// <summary>
/// One of the sixteen colours Word offers for highlighting and for the older colour properties, with the values of <c>WdColorIndex</c>.
/// </summary>
public enum ColorIndex
{
    ByAuthor = -1,
    Auto = 0,
    Black = 1,
    Blue = 2,
    Turquoise = 3,
    BrightGreen = 4,
    Pink = 5,
    Red = 6,
    Yellow = 7,
    White = 8,
    DarkBlue = 9,
    Teal = 10,
    Green = 11,
    Violet = 12,
    DarkRed = 13,
    DarkYellow = 14,
    Gray50 = 15,
    Gray25 = 16
}

/// <summary>
/// The functions of VBA that make colours.
/// </summary>
public static class ColorHelper
{
    /// <summary>
    /// The value of the VBA function <c>RGB</c>: <c>red + 256 * green + 65536 * blue</c>, which is what Word's colour properties hold.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">A component is not from 0 to 255.</exception>
    public static int RGB(int red, int green, int blue)
    {
        return WordColor.FromRgb(red, green, blue).Value;
    }
}

/// <summary>
/// A colour, as Word's colour properties (<c>Font.Color</c>, <c>Border.Color</c>, <c>Shading.BackgroundPatternColor</c>) hold it: a
/// colour or the automatic colour.
/// </summary>
/// <remarks>
/// <para>A colour converts implicitly from a six digit hexadecimal red, green and blue string such as <c>"FF0000"</c> (or <c>"auto"</c>), from a
/// number as <see cref="ColorHelper.RGB"/> makes it, and from a <see cref="NamedColor"/>, so that all of
/// <c>font.Color = "FF0000"</c>, <c>font.Color = ColorHelper.RGB(255, 0, 0)</c> and <c>font.Color = NamedColor.Red</c> work. The default
/// value is the automatic colour.</para>
/// </remarks>
public readonly struct WordColor : IEquatable<WordColor>
{
    private const int AutomaticValue = -16777216;

    //Blue, green and red packed as Word does; only meaningful when _specified.
    private readonly int _bgr;
    private readonly bool _specified;

    private WordColor(int bgr)
    {
        _bgr = bgr;
        _specified = true;
    }

    /// <summary>Makes a colour from a six digit hexadecimal red, green and blue string, or <c>auto</c>.</summary>
    /// <exception cref="ArgumentException">The text is neither six hexadecimal digits nor <c>auto</c>.</exception>
    public WordColor(string hex)
    {
        Guard.ThrowIfNull(hex);

        if (String.Equals(hex, "auto", StringComparison.OrdinalIgnoreCase))
        {
            _bgr = 0;
            _specified = false;
            return;
        }

        int rgb;
        if (hex.Length != 6 || !Int32.TryParse(hex, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out rgb))
        {
            throw new ArgumentException("The colour must be six hexadecimal digits (such as FF0000) or 'auto'.", nameof(hex));
        }

        _bgr = ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF);
        _specified = true;
    }

    //Reads a colour as a document stores it. Text that is not a colour (a file can hold anything) reads as the automatic colour.
    internal static WordColor FromStored(string stored)
    {
        if (stored == null || (stored.Length != 6 && !String.Equals(stored, "auto", StringComparison.OrdinalIgnoreCase)))
        {
            return default(WordColor);
        }

        try
        {
            return new WordColor(stored);
        }
        catch (ArgumentException)
        {
            return default(WordColor);
        }
    }

    /// <summary>The automatic colour.</summary>
    public static WordColor Automatic
    {
        get { return default(WordColor); }
    }

    /// <summary>Makes a colour from its components.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A component is not from 0 to 255.</exception>
    public static WordColor FromRgb(int red, int green, int blue)
    {
        CheckComponent(red, nameof(red));
        CheckComponent(green, nameof(green));
        CheckComponent(blue, nameof(blue));
        return new WordColor(red | (green << 8) | (blue << 16));
    }

    /// <summary>Makes a colour from the number Word holds (<c>WdColor</c>, or the result of <see cref="ColorHelper.RGB"/>).</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is neither a colour (0 to 16777215) nor the automatic colour.</exception>
    public static WordColor FromValue(int value)
    {
        if (value == AutomaticValue)
        {
            return Automatic;
        }

        if (value < 0 || value > 0xFFFFFF)
        {
            throw new ArgumentOutOfRangeException(nameof(value), "The value must be a colour from 0 to 16777215 or the automatic colour.");
        }
        return new WordColor(value);
    }

    /// <summary>Makes a colour from one of the colours of Word's colour index.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The index is not one of the sixteen colours or <see cref="ColorIndex.Auto"/>.</exception>
    public static WordColor FromColorIndex(ColorIndex index)
    {
        switch (index)
        {
            case ColorIndex.Auto: return Automatic;
            case ColorIndex.Black: return FromValue((int)NamedColor.Black);
            case ColorIndex.Blue: return FromValue((int)NamedColor.Blue);
            case ColorIndex.Turquoise: return FromValue((int)NamedColor.Turquoise);
            case ColorIndex.BrightGreen: return FromValue((int)NamedColor.BrightGreen);
            case ColorIndex.Pink: return FromValue((int)NamedColor.Pink);
            case ColorIndex.Red: return FromValue((int)NamedColor.Red);
            case ColorIndex.Yellow: return FromValue((int)NamedColor.Yellow);
            case ColorIndex.White: return FromValue((int)NamedColor.White);
            case ColorIndex.DarkBlue: return FromValue((int)NamedColor.DarkBlue);
            case ColorIndex.Teal: return FromValue((int)NamedColor.Teal);
            case ColorIndex.Green: return FromValue((int)NamedColor.Green);
            case ColorIndex.Violet: return FromValue((int)NamedColor.Violet);
            case ColorIndex.DarkRed: return FromValue((int)NamedColor.DarkRed);
            case ColorIndex.DarkYellow: return FromValue((int)NamedColor.DarkYellow);
            case ColorIndex.Gray50: return FromValue((int)NamedColor.Gray50);
            case ColorIndex.Gray25: return FromValue((int)NamedColor.Gray25);
            default: throw new ArgumentOutOfRangeException(nameof(index));
        }
    }

    /// <summary>Whether this is the automatic colour.</summary>
    public bool IsAutomatic
    {
        get { return !_specified; }
    }

    /// <summary>The number Word holds for the colour (<c>WdColor</c>): <see cref="ColorHelper.RGB"/> of its components, or <c>-16777216</c> for the automatic colour.</summary>
    public int Value
    {
        get { return _specified ? _bgr : AutomaticValue; }
    }

    /// <summary>The red component, 0 for the automatic colour.</summary>
    public int Red
    {
        get { return _bgr & 0xFF; }
    }

    /// <summary>The green component, 0 for the automatic colour.</summary>
    public int Green
    {
        get { return (_bgr >> 8) & 0xFF; }
    }

    /// <summary>The blue component, 0 for the automatic colour.</summary>
    public int Blue
    {
        get { return (_bgr >> 16) & 0xFF; }
    }

    /// <summary>The colour as six uppercase hexadecimal digits (red, green, blue) as Open XML stores it, or <c>auto</c>.</summary>
    public string Hex
    {
        get { return _specified ? Red.ToString("X2", CultureInfo.InvariantCulture) + Green.ToString("X2", CultureInfo.InvariantCulture) + Blue.ToString("X2", CultureInfo.InvariantCulture) : "auto"; }
    }

    /// <summary>The colour of Word's colour index that is nearest to this one (<see cref="ColorIndex.Auto"/> for the automatic colour).</summary>
    public ColorIndex ToColorIndex()
    {
        if (!_specified)
        {
            return ColorIndex.Auto;
        }

        ColorIndex best = ColorIndex.Black;
        long bestDistance = Int64.MaxValue;
        for (int index = (int)ColorIndex.Black; index <= (int)ColorIndex.Gray25; index++)
        {
            WordColor candidate = FromColorIndex((ColorIndex)index);
            long dr = Red - candidate.Red;
            long dg = Green - candidate.Green;
            long db = Blue - candidate.Blue;
            long distance = (dr * dr) + (dg * dg) + (db * db);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = (ColorIndex)index;
            }
        }
        return best;
    }

    private static void CheckComponent(int component, string name)
    {
        if (component < 0 || component > 255)
        {
            throw new ArgumentOutOfRangeException(name, "A colour component must be from 0 to 255.");
        }
    }

    /// <summary>Converts six hexadecimal digits (or <c>auto</c>) to a colour.</summary>
    public static implicit operator WordColor(string hex)
    {
        return new WordColor(hex);
    }

    /// <summary>Converts a number as <see cref="ColorHelper.RGB"/> makes it to a colour.</summary>
    public static implicit operator WordColor(int value)
    {
        return FromValue(value);
    }

    /// <summary>Converts a named colour to a colour.</summary>
    public static implicit operator WordColor(NamedColor color)
    {
        return FromValue((int)color);
    }

    /// <inheritdoc />
    public bool Equals(WordColor other)
    {
        return _specified == other._specified && (!_specified || _bgr == other._bgr);
    }

    /// <inheritdoc />
    public override bool Equals(object obj)
    {
        return obj is WordColor other && Equals(other);
    }

    /// <inheritdoc />
    public override int GetHashCode()
    {
        return _specified ? _bgr : -1;
    }

    /// <summary>Whether two colours are the same.</summary>
    public static bool operator ==(WordColor left, WordColor right)
    {
        return left.Equals(right);
    }

    /// <summary>Whether two colours are different.</summary>
    public static bool operator !=(WordColor left, WordColor right)
    {
        return !left.Equals(right);
    }

    /// <summary>The colour as <see cref="Hex"/>.</summary>
    public override string ToString()
    {
        return Hex;
    }
}
