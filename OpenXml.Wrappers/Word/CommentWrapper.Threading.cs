using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Office2019.Word.Cid;
using DocumentFormat.OpenXml.Office2021.Word.CommentsExt;
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

    internal static string NewParagraphId(MainDocumentPart mainDocumentPart)
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
            part.CommentsEx = CommentParts.NewRoot(new CommentsEx(), "w15", CommentParts.W15);
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
            entry = new CommentEx() { ParaId = paragraphId, Done = OpenXmlExtensions.OffValue() };
            commentsEx.AppendChild(entry);
        }
        return entry;
    }

    private void RemoveThreadInformation()
    {
        CommentParts.Remove(MainPart, _commentRangeStart.GetComment());
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
            entry.Done = value ? OpenXmlExtensions.OnValue() : OpenXmlExtensions.OffValue();
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
        string replyParagraphId = reply.Elements<Paragraph>().Last().ParagraphId.Value;
        GetCommentsEx(true).Elements<CommentEx>().First(e => e.ParaId?.Value == replyParagraphId).ParaIdParent = parentParagraphId;

        //The reply is on the same text as the comment, after the replies that are there already, and its reference follows theirs.
        CommentRangeStart start = new CommentRangeStart() { Id = idText };
        CommentRangeEnd end = new CommentRangeEnd() { Id = idText };
        Run referenceRun = new Run(new RunProperties(new RunStyle() { Val = "CommentReference" }), new CommentReference() { Id = idText });
        previous._commentRangeStart.InsertAfterSelf(start);
        previous.GetReferenceRun().InsertAfterSelf(end);
        end.InsertAfterSelf(referenceRun);

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

/// <summary>
/// The parts that hold what Word records about each comment beside its text: done and threads (<c>commentsExtended</c>), its lasting id
/// (<c>commentsIds</c>), the time in UTC (<c>commentsExtensible</c>) and its author (<c>people</c>). Word writes all of them for every comment.
/// </summary>
internal static class CommentParts
{
    internal const string W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    internal const string W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private const string W16Cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    private const string W16Cex = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
    private const string MarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static readonly Random Randomizer = new Random();

    /// <summary>The root of a part, marked as Word marks it, so that a program that does not know the namespace can ignore it.</summary>
    internal static T NewRoot<T>(T root, string prefix, string namespaceUri) where T : OpenXmlPartRootElement
    {
        root.AddNamespaceDeclaration("mc", MarkupCompatibility);
        root.AddNamespaceDeclaration(prefix, namespaceUri);
        root.MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = prefix };
        return root;
    }

    internal static Comments NewComments()
    {
        return NewRoot(new Comments(), "w14", W14);
    }

    /// <summary>The initials Word would make of a name: the first letter of each word.</summary>
    internal static string GetInitials(string author)
    {
        return string.Concat(author.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(w => w.Substring(0, 1)));
    }

    /// <summary>Gives a new comment its paragraph id and its entry in each of the parts.</summary>
    internal static void Add(MainDocumentPart main, Comment comment)
    {
        string paragraphId = CommentWrapper.NewParagraphId(main);
        comment.Elements<Paragraph>().Last().ParagraphId = paragraphId;
        string durableId = Randomizer.Next(0x10000000, 0x7FFFFFFF).ToString("X8", CultureInfo.InvariantCulture);

        WordprocessingCommentsExPart exPart = main.WordprocessingCommentsExPart ?? main.AddNewPart<WordprocessingCommentsExPart>();
        if (exPart.CommentsEx == null)
        {
            exPart.CommentsEx = NewRoot(new CommentsEx(), "w15", W15);
        }
        exPart.CommentsEx.AppendChild(new CommentEx() { ParaId = paragraphId, Done = OpenXmlExtensions.OffValue() });

        WordprocessingCommentsIdsPart idsPart = main.WordprocessingCommentsIdsPart ?? main.AddNewPart<WordprocessingCommentsIdsPart>();
        if (idsPart.CommentsIds == null)
        {
            idsPart.CommentsIds = NewRoot(new CommentsIds(), "w16cid", W16Cid);
        }
        idsPart.CommentsIds.AppendChild(new CommentId() { ParaId = paragraphId, DurableId = durableId });

        WordCommentsExtensiblePart extensiblePart = main.WordCommentsExtensiblePart ?? main.AddNewPart<WordCommentsExtensiblePart>();
        if (extensiblePart.CommentsExtensible == null)
        {
            extensiblePart.CommentsExtensible = NewRoot(new CommentsExtensible(), "w16cex", W16Cex);
        }
        extensiblePart.CommentsExtensible.AppendChild(new CommentExtensible() { DurableId = durableId, DateUtc = DateTime.UtcNow.Truncate(TimeSpan.TicksPerSecond) });

        AddPerson(main, comment.Author?.Value);
    }

    /// <summary>Takes away the entries of a comment that is going. The author stays listed: Word keeps its people too.</summary>
    internal static void Remove(MainDocumentPart main, Comment comment)
    {
        string paragraphId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
        if (paragraphId == null)
        {
            return;
        }

        main.WordprocessingCommentsExPart?.CommentsEx?.Elements<CommentEx>().FirstOrDefault(e => SameId(e.ParaId?.Value, paragraphId))?.Remove();
        CommentId id = main.WordprocessingCommentsIdsPart?.CommentsIds?.Elements<CommentId>().FirstOrDefault(e => SameId(e.ParaId?.Value, paragraphId));
        if (id != null)
        {
            main.WordCommentsExtensiblePart?.CommentsExtensible?.Elements<CommentExtensible>().FirstOrDefault(e => SameId(e.DurableId?.Value, id.DurableId?.Value))?.Remove();
            id.Remove();
        }
    }

    internal static void DateChanged(MainDocumentPart main, Comment comment, DateTime? date)
    {
        string paragraphId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
        CommentId id = main.WordprocessingCommentsIdsPart?.CommentsIds?.Elements<CommentId>().FirstOrDefault(e => SameId(e.ParaId?.Value, paragraphId));
        CommentExtensible entry = id == null ? null : main.WordCommentsExtensiblePart?.CommentsExtensible?.Elements<CommentExtensible>().FirstOrDefault(e => SameId(e.DurableId?.Value, id.DurableId?.Value));
        if (entry != null)
        {
            entry.DateUtc = date.HasValue ? date.Value.ToUniversalTime() : null;
        }
    }

    /// <summary>Lists the new author, and takes the old one off the list if nothing in the document is by them any more.</summary>
    internal static void AuthorChanged(MainDocumentPart main, string previous, string author)
    {
        AddPerson(main, author);
        if (previous == null || previous == author)
        {
            return;
        }

        bool stillUsed = main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants())
            .Any(e => e.GetAttributes().Any(a => a.LocalName == "author" && a.Value == previous));
        if (!stillUsed)
        {
            main.WordprocessingPeoplePart?.People?.Elements<Person>().FirstOrDefault(p => p.Author?.Value == previous)?.Remove();
        }
    }

    private static void AddPerson(MainDocumentPart main, string author)
    {
        if (author == null)
        {
            return;
        }
        WordprocessingPeoplePart part = main.WordprocessingPeoplePart ?? main.AddNewPart<WordprocessingPeoplePart>();
        if (part.People == null)
        {
            part.People = NewRoot(new People(), "w15", W15);
        }
        if (!part.People.Elements<Person>().Any(p => p.Author?.Value == author))
        {
            part.People.AppendChild(new Person(new PresenceInfo() { ProviderId = "None", UserId = author }) { Author = author });
        }
    }

    private static bool SameId(string first, string second)
    {
        return first != null && string.Equals(first, second, StringComparison.OrdinalIgnoreCase);
    }
}
