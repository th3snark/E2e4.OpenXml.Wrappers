using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Paragraphs: adding them, formatting them and moving between them.
/// </summary>
internal static class ParagraphExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            //The first paragraph exists already. Paragraphs.Add appends one after the last paragraph.
            document.Paragraphs[0].Range.Text = "Centred heading";
            ParagraphWrapper body = document.Paragraphs.Add();
            body.Range.Text = "This paragraph is justified and indented on both sides. It is long enough to run over more than one line so that the alignment and the line spacing can be seen at work in the document.";
            ParagraphWrapper note = document.Paragraphs.Add();
            note.Range.Text = "A note with space above and below it and a double line spacing rule.";

            //Format has the same members as Word's ParagraphFormat. Sizes are in points.
            ParagraphFormatWrapper heading = document.Paragraphs[0].Format;
            heading.Alignment = ParagraphAlignment.Center;
            heading.SpaceAfter = 18;
            heading.KeepWithNext = true;
            document.Paragraphs[0].Range.Font.Size = 20;
            document.Paragraphs[0].Range.Font.Bold = true;

            ParagraphFormatWrapper format = body.Format;
            format.Alignment = ParagraphAlignment.Justify;
            format.LeftIndent = 36;
            format.RightIndent = 36;
            format.FirstLineIndent = 18;
            format.LineSpacingRule = LineSpacingRule.OnePointFive;

            note.Format.SpaceBefore = 12;
            note.Format.SpaceAfter = 12;
            note.Format.LineSpacingRule = LineSpacingRule.DoubleSpacing;
            note.Range.Font.Italic = true;

            //Next and Previous walk the paragraphs in document order, and return null at the ends.
            ParagraphWrapper last = document.Paragraphs[document.Paragraphs.Count - 1];
            for (ParagraphWrapper paragraph = last; paragraph != null; paragraph = paragraph.Previous())
            {
                paragraph.Range.InsertBefore(paragraph.Previous() == null ? "(first) " : "(not first) ");
            }
        }
    }
}
