namespace OpenXml.Wrappers.Word;

public partial class CommentWrapper
{
    /// <summary>
    /// The range over the mark that stands for the comment in the text, after the end of <see cref="Scope"/>, as <c>Comment.Reference</c> in Word.
    /// Its text is empty.
    /// </summary>
    public RangeWrapper Reference
    {
        get
        {
            RangeWrapper scope = Scope;
            return scope.Document.Range(scope.End, scope.End + 1);
        }
    }

    /// <summary>
    /// Deletes the comment and its replies, as <c>Comment.DeleteRecursively</c> in Word. <see cref="Delete"/> does the same, because it takes the
    /// replies with it.
    /// </summary>
    public void DeleteRecursively()
    {
        Delete();
    }
}
