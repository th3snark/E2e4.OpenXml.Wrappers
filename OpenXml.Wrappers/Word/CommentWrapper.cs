using System;
using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class CommentWrapper
{
    CommentRangeStart _commentRangeStart;
    CommentRangeEnd _commentRangeEnd;

    public CommentWrapper(CommentRangeStart commentRangeStart)
    {
        _commentRangeStart = commentRangeStart;
        _commentRangeEnd = commentRangeStart.GetCommentRangeEnd();
    }

    public RangeWrapper Range
    {
        get { return new RangeWrapper(_commentRangeStart, _commentRangeEnd); }
    }

    public string Author
    {
        get { return _commentRangeStart.GetComment().Author?.Value; }
        set { _commentRangeStart.GetComment().Author = value; }
    }

    public string Initials
    {
        get { return _commentRangeStart.GetComment().Initials?.Value; }
        set { _commentRangeStart.GetComment().Initials = value; }
    }

    public string Text
    {
        get { return new RangeWrapper(_commentRangeStart.GetComment()).Text; }
    }

    public int Index
    {
        get { return Convert.ToInt32(_commentRangeStart.Id.Value, CultureInfo.InvariantCulture); }
    }
}
