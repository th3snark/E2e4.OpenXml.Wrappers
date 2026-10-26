using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The comments of a range, as <c>Comments</c> in Word. Replies to a comment are comments too, in the order they are in the document.
/// </summary>
public class CommentWrapperCollection : IEnumerable<CommentWrapper>
{
    private RangeWrapper _rangeWrapper;

    public CommentWrapperCollection(RangeWrapper rangeWrapper)
    {
        _rangeWrapper = rangeWrapper;
    }

    /// <summary>
    /// The range the collection was taken from.
    /// </summary>
    public RangeWrapper Parent
    {
        get { return _rangeWrapper; }
    }

    public CommentWrapper this[int index]
    {
        get
        {
            try
            {
                return new CommentWrapper(GetElementList()[index]);
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
        }
    }

    public int Count { get { return GetElementList().Count; } }

    private IList<CommentRangeStart> GetElementList()
    {
        return _rangeWrapper.StartElement.GetElements<CommentRangeStart>(_rangeWrapper.EndElement).ToList();
    }

    /// <summary>
    /// Adds a comment on the range and returns it, as <c>Comments.Add</c> in Word. The author is the name of the user, which can be changed
    /// afterwards.
    /// </summary>
    /// <remarks>
    /// Supported for a span of characters and for the range of a paragraph, table cell, row, table, bookmark or document. The comment spans the
    /// text of the range; where the range has no text a space is added to anchor it.
    /// </remarks>
    /// <exception cref="NotSupportedException">The range is of a type that cannot take a comment.</exception>
    public CommentWrapper Add(RangeWrapper range, string text)
    {
        Guard.ThrowIfNull(range);

        //Check that the range can be anchored before anything is changed so that no orphan comment is left behind.
        OpenXmlElement start = range.StartElement;
        if (!range.IsSpan && !(start is BookmarkStart || start is Paragraph || start is TableCell || start is TableRow || start is Table || start is Document || start is Body))
        {
            throw new NotSupportedException(string.Format(CultureInfo.InvariantCulture, "Adding a comment to a range starting with the type {0} is not currently supported.", start.GetType().Name));
        }

        MainDocumentPart mainDocumentPart = (start ?? range.AnchorElement).GetDocument().MainDocumentPart;
        Comment comment = CreateComment(mainDocumentPart, text);
        string idText = comment.Id.Value;

        CommentRangeStart commentRangeStart = new CommentRangeStart() { Id = idText };
        CommentRangeEnd commentRangeEnd = new CommentRangeEnd() { Id = idText };
        //The comment reference must always be inside a run.
        Run commentReferenceRun = new Run(new RunProperties(new RunStyle() { Val = "CommentReference" }), new CommentReference() { Id = idText });

        if (range.IsSpan)
        {
            range.PlaceMarkers(commentRangeStart, commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(commentReferenceRun);
        }
        else if (start is BookmarkStart)
        {
            start.InsertBeforeSelf(commentRangeStart);
            range.EndElement.InsertAfterSelf(commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(commentReferenceRun);
        }
        else
        {
            //A paragraph, cell, row, table or the document: the comment spans its first run to its last. If there is no run a single space run
            //is added to the first paragraph to define the range.
            OpenXmlElement container = start is Document document ? document.Body : start;
            if (!container.Descendants<Run>().Any())
            {
                Paragraph firstParagraph = container as Paragraph ?? container.Descendants<Paragraph>().FirstOrDefault();
                if (firstParagraph == null)
                {
                    firstParagraph = new Paragraph();
                    container.AppendChild(firstParagraph);
                }
                firstParagraph.AppendChild(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
            }
            container.Descendants<Run>().First().InsertBeforeSelf(commentRangeStart);
            container.Descendants<Run>().Last().InsertAfterSelf(commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(commentReferenceRun);
        }

        return new CommentWrapper(commentRangeStart);
    }

    /// <summary>
    /// Creates the comment itself (its text, author and date) in the comments part, which is created when the document does not have one.
    /// The caller places the markers and the reference in the document.
    /// </summary>
    internal static Comment CreateComment(MainDocumentPart mainDocumentPart, string text)
    {
        WordprocessingCommentsPart wordprocessingCommentsPart = mainDocumentPart.WordprocessingCommentsPart;
        if (wordprocessingCommentsPart == null)
        {
            wordprocessingCommentsPart = mainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
        }
        Comments comments = wordprocessingCommentsPart.Comments;
        if (comments == null)
        {
            comments = new Comments();
            wordprocessingCommentsPart.Comments = comments;
        }

        //The id of the new comment is one more than the largest existing id (ids are not necessarily contiguous or ordered).
        int id = 0;
        if (comments.Elements<Comment>().Any())
        {
            id = comments.Elements<Comment>().Max(e => Convert.ToInt32(e.Id.Value, CultureInfo.InvariantCulture)) + 1;
        }

        string author = string.IsNullOrEmpty(Environment.UserName) ? "Unknown" : Environment.UserName;
        string initials = author.Substring(0, 1);

        if (string.IsNullOrEmpty(text))
        {
            text = " ";
        }

        Comment comment = new Comment(
                    new Paragraph(
                       new ParagraphProperties(new ParagraphStyleId() { Val = "CommentText" }),
                       new Run(new RunProperties(new RunStyle() { Val = "CommentReference" }),
                               new AnnotationReferenceMark()),
                       new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })))
        {
            Id = id.ToString(CultureInfo.InvariantCulture),
            Date = DateTime.UtcNow.Truncate(TimeSpan.TicksPerSecond),
            Author = author,
            Initials = initials
        };
        comments.AppendChild(comment);
        return comment;
    }

    /// <summary>
    /// Deletes every comment of the collection, with its text, range markers and reference.
    /// </summary>
    public void DeleteAll()
    {
        //The comments are fetched first, since they are deleted inside the iteration.
        foreach (CommentWrapper comment in this.ToList())
        {
            comment.Delete();
        }
    }

    #region IEnumerable Members

    public IEnumerator<CommentWrapper> GetEnumerator()
    {
        //The elements are fetched once for the whole enumeration so that enumerating is linear and so that deleting the
        //current item while enumerating does not disturb the enumeration.
        foreach (CommentRangeStart element in GetElementList())
        {
            yield return new CommentWrapper(element);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
