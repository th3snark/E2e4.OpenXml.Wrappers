using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class CommentWrapperCollection : IEnumerable
{
    private RangeWrapper _rangeWrapper;

    public CommentWrapperCollection(RangeWrapper rangeWrapper)
    {
        _rangeWrapper = rangeWrapper;
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
        return _rangeWrapper.Start.GetElements<CommentRangeStart>(_rangeWrapper.End).ToList();
    }

    public CommentWrapper Add(RangeWrapper range, string text)
    {
        MainDocumentPart mainDocumentPart = _rangeWrapper.Start.GetDocument().MainDocumentPart;

        Comments comments = null;
        if (mainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Any())
        {
            comments = mainDocumentPart.WordprocessingCommentsPart.Comments;
        }
        else
        {
            WordprocessingCommentsPart wordprocessingCommentsPart = mainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            wordprocessingCommentsPart.Comments = new Comments();
            comments = wordprocessingCommentsPart.Comments;
        }

        int id = 0;
        if (comments.HasChildren)
        {
            //Calculatethe id of the new comment by finding the id of the last comment and adding one to it.
            //id = Convert.ToInt32(comments.Descendants<Comment>().Select(e => e.Id.Value).Max()) + 1;
            id = Convert.ToInt32(((Comment)comments.LastChild).Id, CultureInfo.InvariantCulture) + 1;
        }

        string author = Environment.UserName;
        string initials = author.Substring(0, 1);

        if (String.IsNullOrEmpty(text))
        {
            text = " ";
        }

        Comment comment = new Comment(
                    new Paragraph(
                       new ParagraphProperties(new ParagraphStyleId() { Val = "CommentText" }),
                       new Run(new RunProperties(new RunStyle() { Val = "CommentReference" }),
                               new AnnotationReferenceMark()),
                       new Run(new Text(text))))
        {
            Id = id.ToString(CultureInfo.InvariantCulture),
            Date = DateTime.UtcNow.Truncate(TimeSpan.TicksPerSecond),
            Author = author,
            Initials = initials
        };
        comments.AppendChild(comment);

        CommentRangeStart commentRangeStart = new CommentRangeStart() { Id = id.ToString(CultureInfo.InvariantCulture) };
        CommentRangeEnd commentRangeEnd = new CommentRangeEnd() { Id = id.ToString(CultureInfo.InvariantCulture) };
        CommentReference commentReference = new CommentReference() { Id = id.ToString(CultureInfo.InvariantCulture) };

        if (range.Start is BookmarkStart)
        {
            range.Start.InsertBeforeSelf(commentRangeStart);
            range.End.InsertAfterSelf(commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(commentReference);
        }
        else if (range.Start is Paragraph)
        {
            range.Start.InsertBefore(commentRangeStart, range.Start.Elements<Run>().First());
            range.End.InsertAfter(commentRangeEnd, range.Start.Elements<Run>().Last());
            commentRangeEnd.InsertAfterSelf(commentReference);
        }
        if (range.Start is TableCell)
        {
            TableCell tableCell = (TableCell)range.Start;
            Paragraph firstParagraph = tableCell.Elements<Paragraph>().First();
            Paragraph lastParagraph = tableCell.Elements<Paragraph>().Last();
            //If there are no runs in any of the paragraphs in the table cell then add a new single space run to the first paragraph.
            //This will be the text that defines the range for the comment. This is a quick workaround for now since it is actually
            //possible to add a CommentReference without adding a CommentRangeStart and CommentRangeEnd.
            if (!tableCell.Descendants<Run>().Any())
            {
                firstParagraph.AppendChild(new Run(new Text(" ")));
            }
            Run firstRun = firstParagraph.Elements<Run>().First();
            Run lastRun = lastParagraph.Elements<Run>().Last();
            firstRun.InsertBeforeSelf(commentRangeStart);
            lastRun.InsertAfterSelf(commentRangeEnd);
            commentRangeEnd.InsertAfterSelf(
               new Run(
                  new RunProperties(
                     new RunStyle() { Val = "CommentReference" }
                  ),
                  commentReference
               )
            );
        }

        return new CommentWrapper(commentRangeStart);
    }

    #region IEnumerable Members

    public IEnumerator GetEnumerator()
    {
        return new CommentWrapperEnumerator(this);
    }

    #endregion

    #region Nested Private Classes

    private sealed class CommentWrapperEnumerator : IEnumerator, IDisposable
    {
        private CommentWrapperCollection _commentWrapperCollection;
        private int _index;
        private CommentWrapper _current;

        public CommentWrapperEnumerator(CommentWrapperCollection commentWrapperCollection)
        {
            _commentWrapperCollection = commentWrapperCollection;
            _index = -1;
            _current = null;
        }

        #region IEnumerator Members

        public object Current
        {
            get { return _current; }
        }

        public bool MoveNext()
        {
            if (++_index >= _commentWrapperCollection.Count)
                return false;
            else
                _current = _commentWrapperCollection[_index];
            return true;
        }

        public void Reset()
        {
            _current = null;
            _index = -1;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            _commentWrapperCollection = null;
            _current = null;
            _index = -1;
        }

        #endregion
    }

    #endregion
}
