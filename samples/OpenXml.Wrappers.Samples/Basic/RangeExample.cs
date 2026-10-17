using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// A <see cref="RangeWrapper"/> is a piece of the document: all of it, a paragraph, a table, a bookmark and so on. Most editing starts from one.
/// </summary>
internal static class RangeExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            //The whole document. A new document has one empty paragraph, and InsertAfter puts text at the end of the last paragraph.
            RangeWrapper content = document.Range;
            content.InsertAfter("A range is the piece of the document you are working on.");

            //InsertParagraphAfter adds a paragraph at the end of the range and returns it.
            ParagraphWrapper second = content.InsertParagraphAfter();
            second.Range.Text = "In British English the word is spelled colour, and the colour of this paragraph is red.";

            //Find and replace over a whole range. Execute returns whether the text was found.
            bool found = document.Range.Find.Execute("colour", matchCase: false, replaceWith: "color", replace: ReplaceMode.All);

            //Character formatting of the runs in a range.
            document.Paragraphs[0].Range.Font.Bold = true;
            second.Range.Font.Italic = true;
            second.Range.Font.Color = "C00000";

            //InsertBefore puts text at the start of the range.
            document.Paragraphs[0].Range.InsertBefore("Range: ");

            //The text of a range uses Word's markers: \r ends a paragraph.
            string text = document.Paragraphs[1].Range.Text;
            document.Paragraphs.Add().Range.Text = string.Format(CultureInfo.InvariantCulture, "Replaced: {0}. The second paragraph is {1} characters long, including its end marker.", found, text.Length);
        }
    }
}
