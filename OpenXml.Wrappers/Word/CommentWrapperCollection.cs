using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

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
            catch (ArgumentOutOfRangeException ex)
            {
                throw new ArgumentOutOfRangeException(String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index), ex);
            }
        }
    }

    public int Count { get { return GetElementList().Count; } }

    private IList<CommentRangeStart> GetElementList()
    {
        return _rangeWrapper.StartElement.GetElements<CommentRangeStart>(_rangeWrapper.EndElement).ToList();
    }

    public CommentWrapper Add(RangeWrapper range, string text)
    {
        Guard.ThrowIfNull(range);

        //Check that the range can be anchored before anything is changed so that no orphan comment is left behind.
        if (!range.IsSpan && !(range.StartElement is BookmarkStart || range.StartElement is Paragraph || range.StartElement is TableCell))
        {
            throw new NotSupportedException(String.Format(CultureInfo.InvariantCulture, "Adding a comment to a range starting with the type {0} is not currently supported.", range.StartElement.GetType().Name));
        }

        MainDocumentPart mainDocumentPart = _rangeWrapper.StartElement.GetDocument().MainDocumentPart;

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

        string author = String.IsNullOrEmpty(Environment.UserName) ? "Unknown" : Environment.UserName;
        string initials = author.Substring(0, 1);

        if (String.IsNullOrEmpty(text))
        {
            text = " ";
        }

        string idText = id.ToString(CultureInfo.InvariantCulture);
        Comment comment = new Comment(
                    new Paragraph(
                       new ParagraphProperties(new ParagraphStyleId() { Val = "CommentText" }),
                       new Run(new RunProperties(new RunStyle() { Val = "CommentReference" }),
                               new AnnotationReferenceMark()),
                       new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })))
        {
            Id = idText,
            Date = DateTime.UtcNow.Truncate(TimeSpan.TicksPerSecond),
            Author = author,
            Initials = initials
        };
        comments.AppendChild(comment);

        CommentRangeStart commentRangeStart = new CommentRangeStart() { Id = idText };
        CommentRangeEnd commentRangeEnd = new CommentRangeEnd() { Id = idText };
        //The comment reference must always be inside a run.
        Run commentReferenceRun = new Run(new RunProperties(new RunStyle() { Val = "CommentReference" }), new CommentReference() { Id = idText });

        if (range.IsSpan)
        {
            range.PlaceMarkers(commentRangeStart, commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(commentReferenceRun);
        }
        else if (range.StartElement is BookmarkStart)
        {
            range.StartElement.InsertBeforeSelf(commentRangeStart);
            range.EndElement.InsertAfterSelf(commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(commentReferenceRun);
        }
        else if (range.StartElement is Paragraph)
        {
            Paragraph paragraph = (Paragraph)range.StartElement;
            //If the paragraph has no runs then add a single space run to define the range for the comment.
            if (!paragraph.Elements<Run>().Any())
            {
                paragraph.AppendChild(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
            }
            paragraph.Elements<Run>().First().InsertBeforeSelf(commentRangeStart);
            paragraph.Elements<Run>().Last().InsertAfterSelf(commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(commentReferenceRun);
        }
        else
        {
            TableCell tableCell = (TableCell)range.StartElement;
            Paragraph firstParagraph = tableCell.Elements<Paragraph>().First();
            //If there are no runs in any of the paragraphs in the table cell then add a new single space run to the first paragraph.
            //This will be the text that defines the range for the comment. This is a quick workaround for now since it is actually
            //possible to add a CommentReference without adding a CommentRangeStart and CommentRangeEnd.
            if (!tableCell.Descendants<Run>().Any())
            {
                firstParagraph.AppendChild(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
            }
            Run firstRun = tableCell.Descendants<Run>().First();
            Run lastRun = tableCell.Descendants<Run>().Last();
            firstRun.InsertBeforeSelf(commentRangeStart);
            lastRun.InsertAfterSelf(commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(commentReferenceRun);
        }

        return new CommentWrapper(commentRangeStart);
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
