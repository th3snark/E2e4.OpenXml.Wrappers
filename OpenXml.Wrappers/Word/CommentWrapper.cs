using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class CommentWrapper : System.IEquatable<CommentWrapper>
{
    CommentRangeStart _commentRangeStart;
    CommentRangeEnd _commentRangeEnd;

    public CommentWrapper(CommentRangeStart commentRangeStart)
    {
        _commentRangeStart = commentRangeStart;
        _commentRangeEnd = commentRangeStart.GetCommentRangeEnd();
    }

    /// <summary>
    /// The range of the document the comment is attached to (Word's <c>Comment.Scope</c>).
    /// </summary>
    public RangeWrapper Scope
    {
        get { return new RangeWrapper(_commentRangeStart, _commentRangeEnd); }
    }

    /// <summary>
    /// The range over the text of the comment itself (Word's <c>Comment.Range</c>), as opposed to <see cref="Scope"/>.
    /// </summary>
    public RangeWrapper Range
    {
        get
        {
            Comment comment = _commentRangeStart.GetComment();
            Paragraph first = comment.Elements<Paragraph>().FirstOrDefault();
            Paragraph last = comment.Elements<Paragraph>().LastOrDefault();
            return first == null ? new RangeWrapper(comment) : new RangeWrapper(first, last);
        }
    }

    /// <summary>
    /// A range over the element that contains the start of the comment range (usually its paragraph).
    /// </summary>
    public RangeWrapper Parent
    {
        get { return new RangeWrapper(_commentRangeStart.Parent); }
    }

    public string Author
    {
        get { return _commentRangeStart.GetComment().Author?.Value; }
        set { _commentRangeStart.GetComment().Author = value; }
    }

    public string Initial
    {
        get { return _commentRangeStart.GetComment().Initials?.Value; }
        set { _commentRangeStart.GetComment().Initials = value; }
    }

    /// <summary>
    /// The text of the comment, with <c>\r</c> between paragraphs. Setting it replaces the body of the comment, so any formatting
    /// in the body is lost; null is the same as the empty string.
    /// </summary>
    public string Text
    {
        get { return new RangeWrapper(_commentRangeStart.GetComment()).Text; }
        set
        {
            Comment comment = _commentRangeStart.GetComment();
            string[] lines = (value ?? string.Empty).Replace("\r\n", "\r", StringComparison.Ordinal).Replace('\n', '\r').Split('\r');
            //A single trailing paragraph mark ends the last paragraph rather than starting an empty one.
            if (lines.Length > 1 && lines[lines.Length - 1].Length == 0)
            {
                lines = lines.Take(lines.Length - 1).ToArray();
            }

            string paragraphId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
            comment.RemoveAllChildren();
            for (int i = 0; i < lines.Length; i++)
            {
                Paragraph paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId() { Val = "CommentText" }));
                if (i == 0)
                {
                    paragraph.AppendChild(new Run(new RunProperties(new RunStyle() { Val = "CommentReference" }), new AnnotationReferenceMark()));
                }
                string line = lines[i];
                if (line.Length > 0)
                {
                    //SetText removes the runs, so the text is added to a copy and the runs moved across.
                    Paragraph textParagraph = new Paragraph();
                    textParagraph.SetText(line);
                    foreach (Run run in textParagraph.Elements<Run>().ToList())
                    {
                        paragraph.AppendChild(run.CloneNode(true));
                    }
                }
                comment.AppendChild(paragraph);
            }

            //The id that threads (replies and done) refer to belongs to the last paragraph, so the new one keeps it.
            if (paragraphId != null)
            {
                comment.Elements<Paragraph>().Last().ParagraphId = paragraphId;
            }
        }
    }

    /// <summary>
    /// The date and time (UTC) the comment was made, or null if it has none.
    /// </summary>
    public DateTime? Date
    {
        get
        {
            DateTimeValue date = _commentRangeStart.GetComment().Date;
            return date != null && date.HasValue ? date.Value : null;
        }
        set { _commentRangeStart.GetComment().Date = value.HasValue ? value.Value.ToUniversalTime() : null; }
    }

    /// <summary>
    /// Deletes the comment: its text, the range markers and the reference in the document.
    /// </summary>
    public void Delete()
    {
        //The replies go with the comment, as in Word.
        foreach (CommentWrapper reply in Replies.ToList())
        {
            reply.Delete();
        }
        RemoveThreadInformation();

        Comment comment = _commentRangeStart.GetComment();
        string id = _commentRangeStart.Id.Value;
        OpenXmlElement root = _commentRangeStart.GetRoot();

        foreach (CommentReference reference in root.Descendants<CommentReference>().Where(e => e.Id.Value == id).ToList())
        {
            OpenXmlElement parent = reference.Parent;
            reference.Remove();
            if (parent is Run && !parent.Elements().Any(e => !(e is RunProperties)))
            {
                parent.Remove();
            }
        }
        _commentRangeEnd?.Remove();
        _commentRangeStart.Remove();
        comment.Remove();
    }

    /// <summary>
    /// The 0-based position of the comment among the comments of the document (<c>Comment.Index</c> in Word, which counts from 1). The id of
    /// the comment in the file is a different number.
    /// </summary>
    public int Index
    {
        get { return _commentRangeStart.GetRoot().Descendants<CommentRangeStart>().ToList().IndexOf(_commentRangeStart); }
    }

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as CommentWrapper);
    }

    public bool Equals(CommentWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_commentRangeStart, other._commentRangeStart);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_commentRangeStart);
    }

    public static bool operator ==(CommentWrapper left, CommentWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(CommentWrapper left, CommentWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
