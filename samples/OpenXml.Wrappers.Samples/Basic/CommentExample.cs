using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Comments: adding them to a paragraph or bookmark, changing them and deleting them.
/// </summary>
internal static class CommentExample
{
    public static void Create(string path)
    {
        //The last argument asks for a comments part, which a document needs before it can hold comments.
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, true))
        {
            document.Paragraphs[0].Range.Text = "The system will be available 99.99% of the time.";
            document.Paragraphs.Add().Range.Text = "Support is provided during office hours.";
            document.Paragraphs.Add().Range.Text = "This sentence is out of date and will be removed.";
            document.Bookmarks.Add("hours", document.Paragraphs[1].Range);

            //A comment on a paragraph.
            CommentWrapper first = document.Comments.Add(document.Paragraphs[0].Range, "Can we really promise this? Please check with operations.");
            first.Author = "Reviewer";
            first.Initial = "RV";

            //A comment on a bookmark, with the text and date changed after it was added.
            CommentWrapper second = document.Comments.Add(document.Bookmarks["hours"].Range, "Which hours?");
            second.Author = "Reviewer";
            second.Initial = "RV";
            second.Text = "Which hours, and in which time zone?";
            second.Date = new DateTime(2026, 3, 14, 9, 30, 0, DateTimeKind.Utc);

            //A comment that is not wanted any more is deleted together with its markers.
            CommentWrapper obsolete = document.Comments.Add(document.Paragraphs[2].Range, "Delete this later.");
            obsolete.Delete();

            document.Paragraphs.Add().Range.Text = string.Format(CultureInfo.InvariantCulture, "This document has {0} comments.", document.Comments.Count);
        }
    }
}
