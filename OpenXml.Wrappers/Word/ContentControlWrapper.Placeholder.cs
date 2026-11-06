using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// An entry of the glossary document that holds the placeholder text of a content control, as <c>BuildingBlock</c> in Word. Only the name and
/// the text of the entry are modelled.
/// </summary>
public class BuildingBlockWrapper
{
    internal BuildingBlockWrapper(string name, string value)
    {
        Name = name;
        Value = value;
    }

    /// <summary>The name of the entry, as <c>BuildingBlock.Name</c> in Word.</summary>
    public string Name { get; }

    /// <summary>The text of the entry, as <c>BuildingBlock.Value</c> in Word.</summary>
    public string Value { get; }
}

public partial class ContentControlWrapper
{
    //Makes the control show its placeholder again: the text of its entry in the placeholder style, and the mark that says so.
    internal void ShowPlaceholder()
    {
        BuildingBlockWrapper entry = PlaceholderText;
        SdtContentRun content = (_sdt as SdtRun)?.SdtContentRun;
        if (entry == null || content == null)
        {
            return;
        }
        content.RemoveAllChildren();
        content.AppendChild(new Run(new RunProperties(new RunStyle() { Val = ContentControlParts.PlaceholderStyle }), OpenXmlExtensions.NewText(entry.Value)));
        SdtProperties properties = GetProperties(true);
        if (properties.GetFirstChild<ShowingPlaceholder>() == null)
        {
            Place(properties, new ShowingPlaceholder());
        }
    }

    /// <summary>
    /// Word shows the placeholder of a control again as soon as the control is empty. This looks at the controls of a story after text was
    /// changed or deleted in it.
    /// </summary>
    internal static void RestoreEmptyPlaceholders(DocumentFormat.OpenXml.OpenXmlElement story)
    {
        foreach (SdtRun sdt in story.Descendants<SdtRun>().ToList())
        {
            SdtProperties properties = sdt.SdtProperties;
            if (properties != null && properties.GetFirstChild<ShowingPlaceholder>() != null)
            {
                //Text typed over the placeholder replaces it: the mark goes, and so does the style of the placeholder text.
                string shown = string.Concat(sdt.SdtContentRun?.Descendants<Text>().Select(t => t.Text) ?? Enumerable.Empty<string>());
                if (shown.Length == 0)
                {
                    new ContentControlWrapper(sdt).ShowPlaceholder();
                }
                else if (shown != new ContentControlWrapper(sdt).PlaceholderText?.Value && sdt.SdtContentRun != null)
                {
                    properties.RemoveAllChildren<ShowingPlaceholder>();
                    foreach (RunProperties run in sdt.SdtContentRun.Descendants<Run>().Select(r => r.RunProperties).Where(r => r?.RunStyle?.Val?.Value == ContentControlParts.PlaceholderStyle).ToList())
                    {
                        run.RunStyle = null;
                        if (!run.HasChildren)
                        {
                            run.Remove();
                        }
                    }
                }
                continue;
            }
            if (properties == null || properties.GetFirstChild<SdtPlaceholder>() == null)
            {
                continue;
            }
            SdtContentRun content = sdt.SdtContentRun;
            bool empty = content == null || (!content.Descendants<Text>().Any(t => t.Text.Length > 0)
                && !content.Descendants<TabChar>().Any() && !content.Descendants<Break>().Any() && !content.Descendants<Drawing>().Any()
                && !content.Descendants<SymbolChar>().Any() && !content.Descendants<FieldChar>().Any());
            if (empty)
            {
                new ContentControlWrapper(sdt).ShowPlaceholder();
            }
        }
    }

    /// <summary>The entry that holds the placeholder text of the control, as <c>ContentControl.PlaceholderText</c> in Word; null if there is none.</summary>
    public BuildingBlockWrapper PlaceholderText
    {
        get
        {
            string name = GetProperties(false)?.GetFirstChild<SdtPlaceholder>()?.GetFirstChild<DocPartReference>()?.Val?.Value;
            MainDocumentPart main = StyleHelper.GetMainDocumentPart(_sdt);
            DocPart part = name == null ? null : main?.GlossaryDocumentPart?.GlossaryDocument?.DocParts?.Elements<DocPart>()
                .FirstOrDefault(p => p.DocPartProperties?.DocPartName?.Val?.Value == name);
            return part == null ? null : new BuildingBlockWrapper(name, string.Concat(part.DocPartBody.Descendants<Text>().Select(t => t.Text)));
        }
    }

    /// <summary>
    /// Gives the control the text it shows while it is empty, as <c>ContentControl.SetPlaceholderText</c> in Word: a new entry in the glossary
    /// document, which replaces the one the control had. A control that is showing its placeholder shows the new text; one that holds text keeps it.
    /// </summary>
    /// <param name="buildingBlock">An entry to take the text from; not supported.</param>
    /// <param name="range">A range to take the text from (only its text is kept).</param>
    /// <param name="text">The text.</param>
    /// <exception cref="NotSupportedException"><paramref name="buildingBlock"/> is given.</exception>
    /// <exception cref="ArgumentException">Neither a range nor text is given.</exception>
    public void SetPlaceholderText(BuildingBlockWrapper buildingBlock = null, RangeWrapper range = null, string text = null)
    {
        if (buildingBlock != null)
        {
            throw new NotSupportedException("Building blocks are not supported.");
        }
        string value = range != null ? range.Text : text;
        if (value == null)
        {
            throw new ArgumentException("A range or text must be given.", nameof(text));
        }
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(_sdt) ?? throw new NotSupportedException("A control can be given placeholder text only in a document.");
        BuiltInStyles.GetOrCreate(main, ContentControlParts.PlaceholderStyle);

        string name = ContentControlParts.AddPlaceholder(main, value);
        SdtProperties properties = GetProperties(true);
        properties.RemoveAllChildren<SdtPlaceholder>();
        Place(properties, new SdtPlaceholder(new DocPartReference() { Val = name }));

        if (ShowingPlaceholderText)
        {
            SdtContentRun content = (_sdt as SdtRun)?.SdtContentRun;
            if (content != null)
            {
                content.RemoveAllChildren();
                content.AppendChild(new Run(new RunProperties(new RunStyle() { Val = ContentControlParts.PlaceholderStyle }), OpenXmlExtensions.NewText(value)));
            }
        }
        ContentControlParts.RemoveUnusedPlaceholders(main);
    }
}
