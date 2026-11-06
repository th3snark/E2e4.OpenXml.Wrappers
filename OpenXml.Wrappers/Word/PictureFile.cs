using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What the header of a picture file says about it: its kind, its size in pixels and the size it has at its own resolution. The formats
/// are the ones Word inserts as pictures that this library can measure: PNG, JPEG, GIF and BMP.
/// </summary>
internal sealed class PictureFile
{
    private const double DefaultDpi = 96;
    private const long EmuPerInch = 914400;
    private const long EmuPerMetre = 36000000;

    private PictureFile(byte[] bytes, string contentType, int widthPixels, int heightPixels, long widthEmu, long heightEmu)
    {
        Bytes = bytes;
        ContentType = contentType;
        WidthPixels = widthPixels;
        HeightPixels = heightPixels;
        WidthEmu = widthEmu;
        HeightEmu = heightEmu;
    }

    internal byte[] Bytes { get; }

    internal string ContentType { get; }

    internal int WidthPixels { get; }

    internal int HeightPixels { get; }

    /// <summary>The width the picture has at its own resolution, in EMU (English Metric Units, 914400 to the inch).</summary>
    internal long WidthEmu { get; }

    /// <summary>The height the picture has at its own resolution, in EMU.</summary>
    internal long HeightEmu { get; }

    private static int BigEndian32(byte[] b, int at)
    {
        return (b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3];
    }

    private static int BigEndian16(byte[] b, int at)
    {
        return (b[at] << 8) | b[at + 1];
    }

    private static int LittleEndian32(byte[] b, int at)
    {
        return b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24);
    }

    private static int LittleEndian16(byte[] b, int at)
    {
        return b[at] | (b[at + 1] << 8);
    }

    private static long FromDpi(int pixels, double dpi)
    {
        return (long)Math.Round(pixels * EmuPerInch / (dpi > 0 ? dpi : DefaultDpi));
    }

    private static long FromPixelsPerMetre(int pixels, int perMetre)
    {
        return perMetre > 0 ? (long)Math.Round((double)pixels * EmuPerMetre / perMetre) : FromDpi(pixels, DefaultDpi);
    }

    /// <summary>Reads the header of the bytes.</summary>
    /// <exception cref="ArgumentException">The bytes are not a PNG, JPEG, GIF or BMP picture.</exception>
    internal static PictureFile Read(byte[] bytes)
    {
        Guard.ThrowIfNull(bytes);
        if (bytes.Length > 24 && bytes[0] == 0x89 && bytes[1] == 'P' && bytes[2] == 'N' && bytes[3] == 'G')
        {
            return ReadPng(bytes);
        }
        if (bytes.Length > 4 && bytes[0] == 0xFF && bytes[1] == 0xD8)
        {
            return ReadJpeg(bytes);
        }
        if (bytes.Length > 10 && bytes[0] == 'G' && bytes[1] == 'I' && bytes[2] == 'F' && bytes[3] == '8')
        {
            int width = LittleEndian16(bytes, 6);
            int height = LittleEndian16(bytes, 8);
            return new PictureFile(bytes, ImagePartContentTypes.Gif, width, height, FromDpi(width, DefaultDpi), FromDpi(height, DefaultDpi));
        }
        if (bytes.Length > 46 && bytes[0] == 'B' && bytes[1] == 'M')
        {
            int width = LittleEndian32(bytes, 18);
            int height = Math.Abs(LittleEndian32(bytes, 22));
            return new PictureFile(bytes, ImagePartContentTypes.Bmp, width, height,
                FromPixelsPerMetre(width, LittleEndian32(bytes, 38)), FromPixelsPerMetre(height, LittleEndian32(bytes, 42)));
        }
        throw new ArgumentException("The file is not a PNG, JPEG, GIF or BMP picture.", nameof(bytes));
    }

    private static PictureFile ReadPng(byte[] bytes)
    {
        int width = BigEndian32(bytes, 16);
        int height = BigEndian32(bytes, 20);
        int perMetreX = 0;
        int perMetreY = 0;
        int at = 8;
        while (at + 12 <= bytes.Length)
        {
            int length = BigEndian32(bytes, at);
            string type = new string(new[] { (char)bytes[at + 4], (char)bytes[at + 5], (char)bytes[at + 6], (char)bytes[at + 7] });
            if (type == "pHYs" && at + 8 + 9 <= bytes.Length && bytes[at + 16] == 1)
            {
                perMetreX = BigEndian32(bytes, at + 8);
                perMetreY = BigEndian32(bytes, at + 12);
            }
            if (type == "IDAT" || type == "IEND" || length < 0)
            {
                break;
            }
            at += 12 + length;
        }
        return new PictureFile(bytes, ImagePartContentTypes.Png, width, height, FromPixelsPerMetre(width, perMetreX), FromPixelsPerMetre(height, perMetreY));
    }

    private static PictureFile ReadJpeg(byte[] bytes)
    {
        double dpiX = DefaultDpi;
        double dpiY = DefaultDpi;
        int at = 2;
        while (at + 4 <= bytes.Length)
        {
            if (bytes[at] != 0xFF)
            {
                at++;
                continue;
            }
            int marker = bytes[at + 1];
            if (marker == 0xFF)
            {
                at++;
                continue;
            }
            if (marker == 0xD8 || (marker >= 0xD0 && marker <= 0xD7) || marker == 0x01)
            {
                at += 2;
                continue;
            }
            int length = BigEndian16(bytes, at + 2);
            if (marker == 0xE0 && length >= 16 && at + 14 <= bytes.Length && bytes[at + 4] == 'J' && bytes[at + 5] == 'F' && bytes[at + 6] == 'I' && bytes[at + 7] == 'F')
            {
                int units = bytes[at + 11];
                int x = BigEndian16(bytes, at + 12);
                int y = BigEndian16(bytes, at + 14);
                if (units == 1 && x > 0 && y > 0)
                {
                    dpiX = x;
                    dpiY = y;
                }
                else if (units == 2 && x > 0 && y > 0)
                {
                    dpiX = x * 2.54;
                    dpiY = y * 2.54;
                }
            }
            bool isFrame = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
            if (isFrame && at + 9 <= bytes.Length)
            {
                int height = BigEndian16(bytes, at + 5);
                int width = BigEndian16(bytes, at + 7);
                return new PictureFile(bytes, ImagePartContentTypes.Jpeg, width, height, FromDpi(width, dpiX), FromDpi(height, dpiY));
            }
            at += 2 + length;
        }
        throw new ArgumentException("The JPEG picture has no frame header.", nameof(bytes));
    }

    /// <summary>
    /// Takes the picture file or link out of the package when no picture in the story uses it any more.
    /// </summary>
    internal static void Release(OpenXmlElement root, OpenXmlPartContainer container, string embed, string link)
    {
        if (container == null)
        {
            return;
        }
        System.Collections.Generic.List<DocumentFormat.OpenXml.Drawing.Blip> remaining = new System.Collections.Generic.List<DocumentFormat.OpenXml.Drawing.Blip>(root.Descendants<DocumentFormat.OpenXml.Drawing.Blip>());
        if (embed != null && !System.Linq.Enumerable.Any(remaining, b => b.Embed?.Value == embed))
        {
            OpenXmlPart part = System.Linq.Enumerable.FirstOrDefault(System.Linq.Enumerable.Select(System.Linq.Enumerable.Where(container.Parts, p => p.RelationshipId == embed), p => p.OpenXmlPart));
            if (part != null)
            {
                container.DeletePart(part);
            }
        }
        if (link != null && !System.Linq.Enumerable.Any(remaining, b => b.Link?.Value == link))
        {
            container.DeleteExternalRelationship(link);
        }
    }

    /// <summary>The kind of image part that holds a picture with the content type.</summary>
    internal static ImagePart AddPart(OpenXmlPartContainer container, string contentType)
    {
        switch (container)
        {
            case MainDocumentPart main:
                return main.AddImagePart(contentType);
            case HeaderPart header:
                return header.AddImagePart(contentType);
            case FooterPart footer:
                return footer.AddImagePart(contentType);
            default:
                throw new NotSupportedException("A picture can be inserted only in the main text, a header or a footer.");
        }
    }
}

/// <summary>The content types of the pictures this library measures.</summary>
internal static class ImagePartContentTypes
{
    internal const string Png = "image/png";
    internal const string Jpeg = "image/jpeg";
    internal const string Gif = "image/gif";
    internal const string Bmp = "image/bmp";
}
