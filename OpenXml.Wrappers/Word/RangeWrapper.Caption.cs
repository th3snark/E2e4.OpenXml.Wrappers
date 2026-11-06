using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Where a caption goes in relation to what it is for, as <c>WdCaptionPosition</c> in Word (with the same values).
/// </summary>
public enum CaptionPosition
{
    /// <summary>Above the paragraphs.</summary>
    Above = 0,

    /// <summary>Below the paragraphs.</summary>
    Below = 1
}

public partial class RangeWrapper
{
    /// <summary>
    /// Adds a caption for the paragraphs of the range, as <c>Range.InsertCaption</c> in Word: a paragraph in the Caption style with the label,
    /// a <c>SEQ</c> field for the number and the title. The numbers of all the captions with the label are worked out again.
    /// </summary>
    /// <param name="label">The label, for example <c>Figure</c>.</param>
    /// <param name="title">The text after the number.</param>
    /// <param name="titleAutoText">An AutoText entry for the title; not supported.</param>
    /// <param name="position">Above or below the paragraphs; below when not given.</param>
    /// <param name="excludeLabel">True to leave the label out and keep the number and title.</param>
    /// <exception cref="ArgumentNullException">The label is null.</exception>
    /// <exception cref="NotSupportedException"><paramref name="titleAutoText"/> is given.</exception>
    public void InsertCaption(string label, string title = "", string titleAutoText = "", CaptionPosition position = CaptionPosition.Below, bool excludeLabel = false)
    {
        Guard.ThrowIfNull(label);
        if (!EnumCompat.IsDefined(position))
        {
            throw new ArgumentOutOfRangeException(nameof(position));
        }
        if (!string.IsNullOrEmpty(titleAutoText))
        {
            throw new NotSupportedException("AutoText titles are not supported.");
        }

        System.Collections.Generic.IList<Paragraph> paragraphs = GetParagraphs();
        if (paragraphs.Count == 0)
        {
            throw new InvalidOperationException("The range has no paragraph to caption.");
        }
        Paragraph first = paragraphs[0];
        Paragraph last = paragraphs[paragraphs.Count - 1];
        DocumentFormat.OpenXml.Packaging.MainDocumentPart main = StyleHelper.GetMainDocumentPart(first);
        Style style = main == null ? null : BuiltInStyles.GetOrCreate(main, "caption");

        Paragraph caption = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = style?.StyleId?.Value ?? "Caption" }));
        if (!excludeLabel)
        {
            caption.AppendChild(new Run(OpenXmlExtensions.NewText(label + " ")));
        }
        caption.AppendChild(new SimpleField(new Run(new RunProperties(new NoProof()), OpenXmlExtensions.NewText("1")))
        {
            Instruction = " SEQ " + label + " \\* ARABIC "
        });
        if (!string.IsNullOrEmpty(title))
        {
            caption.AppendChild(new Run(OpenXmlExtensions.NewText(title)));
        }

        if (position == CaptionPosition.Above)
        {
            caption.ParagraphProperties.KeepNext = new KeepNext();
            first.InsertBeforeSelf(caption);
        }
        else
        {
            last.GetOrAddParagraphProperties().KeepNext = new KeepNext();
            last.InsertAfterSelf(caption);
        }
        NumberCaptions(caption.GetRoot(), label);
    }

    //Works out the numbers of the captions with a label again, in the order of the text.
    private static void NumberCaptions(OpenXmlElement root, string label)
    {
        Regex pattern = new Regex("^\\s*SEQ\\s+\"?" + Regex.Escape(label) + "\"?(\\s|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        int number = 0;
        foreach (SimpleField field in root.Descendants<SimpleField>().Where(f => pattern.IsMatch(f.Instruction?.Value ?? string.Empty)).ToList())
        {
            number++;
            Text text = field.Descendants<Text>().FirstOrDefault();
            if (text != null)
            {
                text.Text = number.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}
