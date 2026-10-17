using System;
using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class CommentWrapper : System.IEquatable<CommentWrapper>
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
