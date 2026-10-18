using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Styles: listing them, adding your own, giving them formatting and applying them.
/// </summary>
internal static class StyleExample
{
    public static void Create(string path)
    {
        //SampleDocument.Create gives the document the usual styles (Normal, Title, Heading 1...).
        using (WordprocessingDocumentWrapper document = SampleDocument.Create(path, "Styles"))
        {
            //Add a paragraph style and a character style, and give them formatting. A style's Font and ParagraphFormat are only what is
            //set in the style itself.
            StyleWrapper callout = document.Styles.Add("Callout", StyleType.Paragraph);
            callout.Font.Italic = true;
            callout.Font.Color = "7F3F00";
            callout.ParagraphFormat.LeftIndent = 28;
            callout.ParagraphFormat.RightIndent = 28;
            callout.ParagraphFormat.SpaceBefore = 8;
            callout.ParagraphFormat.SpaceAfter = 8;

            StyleWrapper keyTerm = document.Styles.Add("Key Term", StyleType.Character);
            keyTerm.Font.Bold = true;
            keyTerm.Font.Color = "C00000";

            //Apply styles by name through Range.Style. A paragraph style goes on whole paragraphs, a character style on the runs.
            document.Paragraphs[0].Range.Text = "Using styles";
            document.Paragraphs[0].Range.Style = "Title";
            document.Paragraphs.Add().Range.Text = "Why styles";
            document.Paragraphs[1].Range.Style = "Heading 1";
            document.Paragraphs.Add().Range.Text = "A style keeps the formatting of the whole document in one place, so changing the style changes every paragraph that uses it.";
            document.Paragraphs.Add().Range.Text = "Tip: change a style, not the paragraphs.";
            document.Paragraphs[3].Range.Style = callout.Name;
            document.Paragraphs.Add().Range.Text = "The words in this sentence are not styled except for one key term.";
            document.Paragraphs[4].Range.Find.Execute("key term", replaceWith: "key term", replace: ReplaceMode.All);
            document.Paragraphs[4].Range.Style = "Key Term";

            //A style that was added and is not wanted can be deleted; built-in styles cannot.
            StyleWrapper unused = document.Styles.Add("Unused", StyleType.Paragraph);
            unused.Delete();

            //List the paragraph styles that the document has.
            IEnumerable<string> names = document.Styles.Where(s => s.Type == StyleType.Paragraph).Select(s => s.Name);
            document.Paragraphs.Add().Range.Text = string.Format(CultureInfo.InvariantCulture, "Paragraph styles in this document: {0}.", string.Join(", ", names));
        }
    }
}
