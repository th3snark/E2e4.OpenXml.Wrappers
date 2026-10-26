using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class CommentWrapper
{
    private static readonly Random Randomizer = new Random();

    private MainDocumentPart MainPart
    {
        get { return _commentRangeStart.GetDocument().MainDocumentPart; }
    }

    //The paragraph id of the last paragraph of the comment, which is what the comment is known by in the file that holds the threads.
    private string GetParagraphId(bool create)
    {
        Paragraph last = _commentRangeStart.GetComment().Elements<Paragraph>().LastOrDefault();
        if (last == null)
        {
            return null;
        }
        if (last.ParagraphId == null && create)
        {
            last.ParagraphId = NewParagraphId(MainPart);
        }
        return last.ParagraphId?.Value;
    }

    private static string NewParagraphId(MainDocumentPart mainDocumentPart)
    {
        HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Paragraph paragraph in mainDocumentPart.Document.Descendants<Paragraph>().Concat(
            mainDocumentPart.WordprocessingCommentsPart?.Comments?.Descendants<Paragraph>() ?? Enumerable.Empty<Paragraph>()))
        {
            if (paragraph.ParagraphId != null)
            {
                used.Add(paragraph.ParagraphId.Value);
            }
        }

        string id;
        do
        {
            //The ids are 8 hexadecimal digits that are less than 80000000.
            id = Randomizer.Next(0x10000000, 0x7FFFFFFF).ToString("X8", CultureInfo.InvariantCulture);
        }
        while (!used.Add(id));
        return id;
    }

    private CommentsEx GetCommentsEx(bool create)
    {
        MainDocumentPart main = MainPart;
        WordprocessingCommentsExPart part = main.WordprocessingCommentsExPart;
        if (part == null)
        {
            if (!create)
            {
                return null;
            }
            part = main.AddNewPart<WordprocessingCommentsExPart>();
        }
        if (part.CommentsEx == null)
        {
            if (!create)
            {
                return null;
            }
            part.CommentsEx = new CommentsEx();
        }
        return part.CommentsEx;
    }

    private CommentEx FindCommentEx(bool create)
    {
        string paragraphId = GetParagraphId(create);
        if (paragraphId == null)
        {
            return null;
        }

        CommentsEx commentsEx = GetCommentsEx(create);
        CommentEx entry = commentsEx?.Elements<CommentEx>().FirstOrDefault(e => string.Equals(e.ParaId?.Value, paragraphId, StringComparison.OrdinalIgnoreCase));
        if (entry == null && create)
        {
            entry = new CommentEx() { ParaId = paragraphId, Done = false };
            commentsEx.AppendChild(entry);
        }
        return entry;
    }

    private void RemoveThreadInformation()
    {
        CommentEx entry = FindCommentEx(false);
        entry?.Remove();
    }

    /// <summary>
    /// Whether the comment has been marked as done (resolved), as <c>Comment.Done</c> in Word.
    /// </summary>
    public bool Done
    {
        get
        {
            CommentEx entry = FindCommentEx(false);
            return entry?.Done != null && entry.Done.HasValue && entry.Done.Value;
        }
        set
        {
            CommentEx entry = FindCommentEx(true);
            entry.Done = value;
        }
    }

    /// <summary>
    /// The comment that this one is a reply to, or null (<c>Comment.Ancestor</c> in Word).
    /// </summary>
    public CommentWrapper Ancestor
    {
        get
        {
            string parentId = FindCommentEx(false)?.ParaIdParent?.Value;
            if (parentId == null)
            {
                return null;
            }

            Comments comments = MainPart.WordprocessingCommentsPart?.Comments;
            Comment parent = comments?.Elements<Comment>().FirstOrDefault(c => string.Equals(c.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value, parentId, StringComparison.OrdinalIgnoreCase));
            if (parent == null)
            {
                return null;
            }
            CommentRangeStart start = _commentRangeStart.GetRoot().Descendants<CommentRangeStart>().FirstOrDefault(s => s.Id?.Value == parent.Id?.Value);
            return start == null ? null : new CommentWrapper(start);
        }
    }

    /// <summary>
    /// The replies to the comment, as <c>Comment.Replies</c> in Word.
    /// </summary>
    public CommentReplyCollection Replies
    {
        get { return new CommentReplyCollection(this); }
    }

    /// <summary>The comments that are replies to this one, in document order.</summary>
    internal IList<CommentWrapper> GetReplies()
    {
        string paragraphId = GetParagraphId(false);
        CommentsEx commentsEx = GetCommentsEx(false);
        if (paragraphId == null || commentsEx == null)
        {
            return new List<CommentWrapper>();
        }

        HashSet<string> replyIds = new HashSet<string>(
            commentsEx.Elements<CommentEx>().Where(e => string.Equals(e.ParaIdParent?.Value, paragraphId, StringComparison.OrdinalIgnoreCase)).Select(e => e.ParaId?.Value),
            StringComparer.OrdinalIgnoreCase);
        if (replyIds.Count == 0)
        {
            return new List<CommentWrapper>();
        }

        Comments comments = MainPart.WordprocessingCommentsPart?.Comments;
        HashSet<string> commentIds = new HashSet<string>(
            comments.Elements<Comment>().Where(c => c.Elements<Paragraph>().LastOrDefault()?.ParagraphId != null && replyIds.Contains(c.Elements<Paragraph>().Last().ParagraphId.Value)).Select(c => c.Id.Value));
        return _commentRangeStart.GetRoot().Descendants<CommentRangeStart>().Where(s => commentIds.Contains(s.Id.Value)).Select(s => new CommentWrapper(s)).ToList();
    }

    internal CommentWrapper AddReply(string text)
    {
        MainDocumentPart main = MainPart;
        string parentParagraphId = GetParagraphId(true);
        FindCommentEx(true);

        IList<CommentWrapper> existing = GetReplies();
        CommentWrapper previous = existing.Count > 0 ? existing[existing.Count - 1] : this;

        Comment reply = CommentWrapperCollection.CreateComment(main, text);
        string idText = reply.Id.Value;
        string replyParagraphId = NewParagraphId(main);
        reply.Elements<Paragraph>().Last().ParagraphId = replyParagraphId;
        GetCommentsEx(true).AppendChild(new CommentEx() { ParaId = replyParagraphId, ParaIdParent = parentParagraphId, Done = false });

        //The reply is on the same text as the comment, after the replies that are there already, and its reference follows theirs.
        CommentRangeStart start = new CommentRangeStart() { Id = idText };
        CommentRangeEnd end = new CommentRangeEnd() { Id = idText };
        Run referenceRun = new Run(new RunProperties(new RunStyle() { Val = "CommentReference" }), new CommentReference() { Id = idText });
        previous._commentRangeEnd.InsertAfterSelf(end);
        previous._commentRangeStart.InsertAfterSelf(start);
        previous.GetReferenceRun().InsertAfterSelf(referenceRun);

        return new CommentWrapper(start);
    }

    private Run GetReferenceRun()
    {
        string id = _commentRangeStart.Id.Value;
        return _commentRangeStart.GetRoot().Descendants<CommentReference>().First(r => r.Id?.Value == id).Ancestors<Run>().First();
    }
}

/// <summary>
/// The replies to a comment, as <c>Comments</c> of <c>Comment.Replies</c> in Word.
/// </summary>
public class CommentReplyCollection : IEnumerable<CommentWrapper>
{
    private readonly CommentWrapper _comment;

    internal CommentReplyCollection(CommentWrapper comment)
    {
        _comment = comment;
    }

    /// <summary>The comment that the replies are to.</summary>
    public CommentWrapper Parent
    {
        get { return _comment; }
    }

    /// <summary>The number of replies.</summary>
    public int Count
    {
        get { return _comment.GetReplies().Count; }
    }

    /// <summary>The reply at the position (in document order).</summary>
    public CommentWrapper this[int index]
    {
        get
        {
            IList<CommentWrapper> replies = _comment.GetReplies();
            if (index < 0 || index >= replies.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return replies[index];
        }
    }

    /// <summary>Adds a reply to the comment with the text and returns it.</summary>
    public CommentWrapper Add(string text)
    {
        return _comment.AddReply(text);
    }

    public IEnumerator<CommentWrapper> GetEnumerator()
    {
        return _comment.GetReplies().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
