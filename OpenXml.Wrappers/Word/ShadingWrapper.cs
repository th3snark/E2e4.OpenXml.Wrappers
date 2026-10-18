using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The pattern of a shading, as <c>WdTextureIndex</c> in Word (with the same values). The percentages are how much of the area the
/// foreground colour covers.
/// </summary>
public enum ShadingTexture
{
    None = 0,
    Percent5 = 50,
    Percent10 = 100,
    Percent12Point5 = 125,
    Percent15 = 150,
    Percent20 = 200,
    Percent25 = 250,
    Percent30 = 300,
    Percent35 = 350,
    Percent37Point5 = 375,
    Percent40 = 400,
    Percent45 = 450,
    Percent50 = 500,
    Percent55 = 550,
    Percent60 = 600,
    Percent62Point5 = 625,
    Percent65 = 650,
    Percent70 = 700,
    Percent75 = 750,
    Percent80 = 800,
    Percent85 = 850,
    Percent87Point5 = 875,
    Percent90 = 900,
    Percent95 = 950,
    Solid = 1000,
    DarkHorizontal = -1,
    DarkVertical = -2,
    DarkDiagonalDown = -3,
    DarkDiagonalUp = -4,
    DarkCross = -5,
    DarkDiagonalCross = -6,
    Horizontal = -7,
    Vertical = -8,
    DiagonalDown = -9,
    DiagonalUp = -10,
    Cross = -11,
    DiagonalCross = -12
}

/// <summary>
/// The shading of a table, row, column or cell, as <c>Shading</c> in Word.
/// </summary>
/// <remarks>
/// A property is null when the cells it covers are shaded differently. Colours are <see cref="WordColor"/> values, as for
/// <see cref="FontWrapper.Color"/>. With the texture <see cref="ShadingTexture.None"/> the background colour fills the area; with a
/// pattern the foreground colour is the colour of the pattern.
/// </remarks>
public class ShadingWrapper
{
    //The pairs of texture and shading pattern.
    private static readonly List<KeyValuePair<ShadingTexture, ShadingPatternValues>> Patterns = new List<KeyValuePair<ShadingTexture, ShadingPatternValues>>
    {
        Pair(ShadingTexture.None, ShadingPatternValues.Clear),
        Pair(ShadingTexture.Solid, ShadingPatternValues.Solid),
        Pair(ShadingTexture.Percent5, ShadingPatternValues.Percent5),
        Pair(ShadingTexture.Percent10, ShadingPatternValues.Percent10),
        Pair(ShadingTexture.Percent12Point5, ShadingPatternValues.Percent12),
        Pair(ShadingTexture.Percent15, ShadingPatternValues.Percent15),
        Pair(ShadingTexture.Percent20, ShadingPatternValues.Percent20),
        Pair(ShadingTexture.Percent25, ShadingPatternValues.Percent25),
        Pair(ShadingTexture.Percent30, ShadingPatternValues.Percent30),
        Pair(ShadingTexture.Percent35, ShadingPatternValues.Percent35),
        Pair(ShadingTexture.Percent37Point5, ShadingPatternValues.Percent37),
        Pair(ShadingTexture.Percent40, ShadingPatternValues.Percent40),
        Pair(ShadingTexture.Percent45, ShadingPatternValues.Percent45),
        Pair(ShadingTexture.Percent50, ShadingPatternValues.Percent50),
        Pair(ShadingTexture.Percent55, ShadingPatternValues.Percent55),
        Pair(ShadingTexture.Percent60, ShadingPatternValues.Percent60),
        Pair(ShadingTexture.Percent62Point5, ShadingPatternValues.Percent62),
        Pair(ShadingTexture.Percent65, ShadingPatternValues.Percent65),
        Pair(ShadingTexture.Percent70, ShadingPatternValues.Percent70),
        Pair(ShadingTexture.Percent75, ShadingPatternValues.Percent75),
        Pair(ShadingTexture.Percent80, ShadingPatternValues.Percent80),
        Pair(ShadingTexture.Percent85, ShadingPatternValues.Percent85),
        Pair(ShadingTexture.Percent87Point5, ShadingPatternValues.Percent87),
        Pair(ShadingTexture.Percent90, ShadingPatternValues.Percent90),
        Pair(ShadingTexture.Percent95, ShadingPatternValues.Percent95),
        Pair(ShadingTexture.DarkHorizontal, ShadingPatternValues.HorizontalStripe),
        Pair(ShadingTexture.DarkVertical, ShadingPatternValues.VerticalStripe),
        Pair(ShadingTexture.DarkDiagonalDown, ShadingPatternValues.ReverseDiagonalStripe),
        Pair(ShadingTexture.DarkDiagonalUp, ShadingPatternValues.DiagonalStripe),
        Pair(ShadingTexture.DarkCross, ShadingPatternValues.HorizontalCross),
        Pair(ShadingTexture.DarkDiagonalCross, ShadingPatternValues.DiagonalCross),
        Pair(ShadingTexture.Horizontal, ShadingPatternValues.ThinHorizontalStripe),
        Pair(ShadingTexture.Vertical, ShadingPatternValues.ThinVerticalStripe),
        Pair(ShadingTexture.DiagonalDown, ShadingPatternValues.ThinReverseDiagonalStripe),
        Pair(ShadingTexture.DiagonalUp, ShadingPatternValues.ThinDiagonalStripe),
        Pair(ShadingTexture.Cross, ShadingPatternValues.ThinHorizontalCross),
        Pair(ShadingTexture.DiagonalCross, ShadingPatternValues.ThinDiagonalCross)
    };

    //Finds the shading elements of everything the shading covers, creating them when asked to. An entry is null when there is none yet.
    private readonly Func<bool, IList<Shading>> _access;

    internal ShadingWrapper(Func<bool, IList<Shading>> access)
    {
        _access = access;
    }

    private static KeyValuePair<ShadingTexture, ShadingPatternValues> Pair(ShadingTexture texture, ShadingPatternValues pattern)
    {
        return new KeyValuePair<ShadingTexture, ShadingPatternValues>(texture, pattern);
    }

    private T Read<T>(Func<Shading, T> read, T none)
    {
        IList<Shading> shadings = _access(false);
        if (shadings.Count == 0)
        {
            return none;
        }

        List<T> values = shadings.Select(s => s == null ? none : read(s)).ToList();
        return values.All(v => EqualityComparer<T>.Default.Equals(v, values[0])) ? values[0] : default(T);
    }

    private void Write(Action<Shading> change)
    {
        foreach (Shading shading in _access(true))
        {
            change(shading);
        }
    }

    /// <summary>The colour of the area; automatic for none (<c>Shading.BackgroundPatternColor</c> in Word). Null when the cells differ.</summary>
    public WordColor? BackgroundPatternColor
    {
        get { return Read<WordColor?>(s => WordColor.FromStored(s.Fill?.Value), WordColor.Automatic); }
        set { Write(s => s.Fill = (value ?? WordColor.Automatic).Hex); }
    }

    /// <summary>The colour of the pattern; automatic for none (<c>Shading.ForegroundPatternColor</c> in Word). Null when the cells differ.</summary>
    public WordColor? ForegroundPatternColor
    {
        get { return Read<WordColor?>(s => WordColor.FromStored(s.Color?.Value), WordColor.Automatic); }
        set { Write(s => s.Color = (value ?? WordColor.Automatic).Hex); }
    }

    /// <summary>The background colour as one of the colours of Word's colour index (<c>Shading.BackgroundPatternColorIndex</c> in Word).</summary>
    public ColorIndex? BackgroundPatternColorIndex
    {
        get { return BackgroundPatternColor?.ToColorIndex(); }
        set { BackgroundPatternColor = value.HasValue ? WordColor.FromColorIndex(value.Value) : (WordColor?)null; }
    }

    /// <summary>The colour of the pattern as one of the colours of Word's colour index (<c>Shading.ForegroundPatternColorIndex</c> in Word).</summary>
    public ColorIndex? ForegroundPatternColorIndex
    {
        get { return ForegroundPatternColor?.ToColorIndex(); }
        set { ForegroundPatternColor = value.HasValue ? WordColor.FromColorIndex(value.Value) : (WordColor?)null; }
    }

    /// <summary>The pattern of the shading (<c>Shading.Texture</c> in Word).</summary>
    public ShadingTexture? Texture
    {
        get { return Read<ShadingTexture?>(s => ToTexture(s.Val), ShadingTexture.None); }
        set
        {
            if (!value.HasValue || !EnumCompat.IsDefined(value.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Write(s => s.Val = ToPattern(value.Value));
        }
    }

    /// <summary>Removes the shading.</summary>
    public void Clear()
    {
        foreach (Shading shading in _access(false).Where(s => s != null))
        {
            shading.Remove();
        }
    }

    private static ShadingTexture ToTexture(EnumValue<ShadingPatternValues> pattern)
    {
        if (pattern == null || !pattern.HasValue)
        {
            return ShadingTexture.None;
        }
        foreach (KeyValuePair<ShadingTexture, ShadingPatternValues> pair in Patterns)
        {
            if (pair.Value == pattern.Value)
            {
                return pair.Key;
            }
        }
        return ShadingTexture.None;
    }

    private static ShadingPatternValues ToPattern(ShadingTexture texture)
    {
        foreach (KeyValuePair<ShadingTexture, ShadingPatternValues> pair in Patterns)
        {
            if (pair.Key == texture)
            {
                return pair.Value;
            }
        }
        throw new ArgumentOutOfRangeException(nameof(texture));
    }
}
