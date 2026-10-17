using System;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper : System.IEquatable<RangeWrapper>
{
    OpenXmlElement _startOpenXmlElement;
    OpenXmlElement _endOpenXmlElement;

    public RangeWrapper(OpenXmlElement startOpenXmlElement)
       : this(startOpenXmlElement, startOpenXmlElement)
    {
    }

    public RangeWrapper(OpenXmlElement startOpenXmlElement, OpenXmlElement endOpenXmlElement)
    {
        _startOpenXmlElement = startOpenXmlElement;
        _endOpenXmlElement = endOpenXmlElement;
    }

    /// <summary>
    /// The element the range starts at.
    /// </summary>
    public OpenXmlElement StartElement
    {
        get { return _startOpenXmlElement; }
    }

    /// <summary>
    /// The element the range ends at.
    /// </summary>
    public OpenXmlElement EndElement
    {
        get { return _endOpenXmlElement; }
    }

    public BookmarkWrapperCollection Bookmarks
    {
        get { return new BookmarkWrapperCollection(ElementRange()); }
    }

    public TableWrapperCollection Tables
    {
        get { return new TableWrapperCollection(ElementRange()); }
    }

    public TableCellWrapperCollection Cells
    {
        get { return new TableCellWrapperCollection(ElementRange()); }
    }

    public TableRowWrapperCollection Rows
    {
        get { return new TableRowWrapperCollection(ElementRange()); }
    }

    public ParagraphWrapperCollection Paragraphs
    {
        get { return new ParagraphWrapperCollection(this); }
    }

    public CommentWrapperCollection Comments
    {
        get { return new CommentWrapperCollection(ElementRange()); }
    }

    public string Text
    {
        get
        {
            if (_isSpan)
            {
                return GetSpanText();
            }

            StringBuilder sb = new StringBuilder();
            if (_startOpenXmlElement is BookmarkStart)
            {
                AppendBookmarkText(sb, (BookmarkStart)_startOpenXmlElement, (BookmarkEnd)_endOpenXmlElement);
            }
            else if (_startOpenXmlElement is CommentRangeStart)
            {
                AppendCommentText(sb, (CommentRangeStart)_startOpenXmlElement, (CommentRangeEnd)_endOpenXmlElement);
            }
            else if (_startOpenXmlElement is Comment)
            {
                AppendCommentText(sb, (Comment)_startOpenXmlElement);
            }
            else if (_startOpenXmlElement is Paragraph)
            {
                Paragraph paragraph = (Paragraph)_startOpenXmlElement;
                AppendParagraphText(sb, paragraph);
            }
            else if (_startOpenXmlElement is TableCell)
            {
                TableCell tableCell = (TableCell)_startOpenXmlElement;
                AppendTableCellText(sb, tableCell);
            }
            else if (_startOpenXmlElement is TableRow)
            {
                TableRow tableRow = (TableRow)_startOpenXmlElement;
                AppendTableRowText(sb, tableRow);
            }
            else if (_startOpenXmlElement is Table)
            {
                Table table = (Table)_startOpenXmlElement;
                AppendTableText(sb, table);
            }
            else if (_startOpenXmlElement is Document || _startOpenXmlElement is Body)
            {
                StoryMap map = StoryMap.Build(StoryMap.FindRoot(_startOpenXmlElement));
                sb.Append(map.GetText(0, map.Length));
            }
            else if (_startOpenXmlElement is Header || _startOpenXmlElement is Footer)
            {
                foreach (OpenXmlElement block in _startOpenXmlElement.ChildElements)
                {
                    if (block is Paragraph)
                    {
                        AppendParagraphText(sb, (Paragraph)block);
                    }
                    else if (block is Table)
                    {
                        AppendTableText(sb, (Table)block);
                    }
                }
            }
            else
            {
                throw new InvalidOperationException(String.Format(CultureInfo.InvariantCulture, "Getting the Text property is not currently supported for the type {0}", _startOpenXmlElement.GetType().Name));
            }
            return sb.ToString();
        }

        set
        {
            string text = value == null ? String.Empty : value;

            if (_isSpan)
            {
                SetSpanText(text);
            }
            else if (_startOpenXmlElement is TableCell)
            {
                //Get the first paragraph in the table cell.
                Paragraph paragraph = _startOpenXmlElement.Elements<Paragraph>().First();

                //Set the paragraph text to the desired value after first removing any end of table cell characters.
                SetParagraphRunText(paragraph, text.TrimEnd(new char[] { '\r', '\a' }));
            }
            else if (_startOpenXmlElement is Paragraph)
            {
                SetParagraphRunText((Paragraph)_startOpenXmlElement, text);
            }
            else if (_startOpenXmlElement is BookmarkStart)
            {
                List<Run> oldRunList = _startOpenXmlElement.GetElements<Run>(_endOpenXmlElement).ToList();
                Paragraph paragraph = _startOpenXmlElement.Ancestors<Paragraph>().First();
                List<Run> newRunList = BuildNewRunListFromText(text, paragraph);

                //Find the target open xml element that we will insert the new runs immediately after.
                OpenXmlElement targetOpenXmlElement = null;
                if (oldRunList.Count == 0)
                {
                    targetOpenXmlElement = _startOpenXmlElement;
                }
                else
                {
                    targetOpenXmlElement = oldRunList.Last();
                }

                //Insert the new runs (in reverse order) immediately after the target open xml element.
                newRunList.Reverse();
                foreach (Run run in newRunList)
                {
                    targetOpenXmlElement.InsertAfterSelf<Run>(run);
                }

                //If necessary remove any old runs from parent.
                if (oldRunList.Count > 0)
                {
                    DeleteAllRunsInList(oldRunList);
                }
            }
            else if (_startOpenXmlElement is Document || _startOpenXmlElement is Body)
            {
                //All the content is replaced by the text: it is emptied (leaving the one paragraph that a body needs) and the text typed into it.
                EmptyBody(GetBody());
                RangeWrapper whole = new RangeWrapper(StoryMap.Build(StoryMap.FindRoot(_startOpenXmlElement)), 0, 0);
                whole.SetSpanText(text.EndsWith('\r') ? text.Substring(0, text.Length - 1) : text);
            }
            else if (_startOpenXmlElement is Header || _startOpenXmlElement is Footer)
            {
                //The content is replaced by one paragraph for each line, so the header or footer ends up with plain paragraphs.
                _startOpenXmlElement.RemoveAllChildren();
                foreach (string line in text.TrimEnd('\r').Split('\r'))
                {
                    Paragraph paragraph = new Paragraph();
                    foreach (Run run in BuildNewRunListFromText(line, paragraph))
                    {
                        paragraph.AppendChild(run);
                    }
                    _startOpenXmlElement.AppendChild(paragraph);
                }
            }
            else
            {
                throw new InvalidOperationException(String.Format(CultureInfo.InvariantCulture, "Setting the Text property is not currently supported for the type {0}", _startOpenXmlElement.GetType().Name));
            }
        }
    }

    private void SetParagraphRunText(Paragraph paragraph, string value)
    {
        List<Run> oldRunList = paragraph.GetElements<Run>().ToList();
        List<Run> newRunList = BuildNewRunListFromText(value, paragraph);
        DeleteAllRunsInList(oldRunList);
        foreach (Run run in newRunList)
        {
            paragraph.AppendChild(run);
        }
    }

    private List<Run> BuildNewRunListFromText(string value, Paragraph paragraph)
    {
        ParagraphMarkRunProperties paragraphMarkRunProperties = null;
        if (paragraph.ParagraphProperties != null)
        {
            paragraphMarkRunProperties = paragraph.ParagraphProperties.ParagraphMarkRunProperties;
        }

        return TextToRunList(value, paragraphMarkRunProperties);
    }

    private List<Run> TextToRunList(string value, ParagraphMarkRunProperties paragraphMarkRunProperties)
    {
        List<Run> runList = new List<Run>();

        StringBuilder sb = new StringBuilder();
        Run run = null;
        foreach (char c in value.Reverse())
        {
            switch (c)
            {
                case '\t':
                    run = new Run(new TabChar());
                    if (sb.Length > 0)
                    {
                        run.AppendChild<Text>(new Text(sb.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                    }
                    PrependRunProperties(run, paragraphMarkRunProperties);
                    runList.Insert(0, run);
                    sb.Length = 0;
                    break;

                case '\f':
                    run = new Run(new Break() { Type = BreakValues.Page });
                    if (sb.Length > 0)
                    {
                        run.AppendChild<Text>(new Text(sb.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                    }
                    PrependRunProperties(run, paragraphMarkRunProperties);
                    runList.Insert(0, run);
                    sb.Length = 0;
                    break;

                case '\n':
                case '\v':
                    run = new Run(new Break());
                    if (sb.Length > 0)
                    {
                        run.AppendChild<Text>(new Text(sb.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                    }
                    PrependRunProperties(run, paragraphMarkRunProperties);
                    runList.Insert(0, run);
                    sb.Length = 0;
                    break;

                default:
                    sb.Insert(0, c);
                    break;
            }
        }
        if (sb.Length > 0)
        {
            run = new Run(new Text(sb.ToString()) { Space = SpaceProcessingModeValues.Preserve });
            PrependRunProperties(run, paragraphMarkRunProperties);

            runList.Insert(0, run);
        }
        return runList;
    }

    private static void PrependRunProperties(Run run, ParagraphMarkRunProperties paragraphMarkRunProperties)
    {
        if (paragraphMarkRunProperties == null)
        {
            return;
        }

        //The run properties of a run are a different type to the paragraph mark run properties. Only the formatting elements
        //are copied since the tracked change elements are only valid on the paragraph mark.
        RunProperties runProperties = new RunProperties(
            paragraphMarkRunProperties.ChildElements
                .Where(e => !(e is Inserted || e is Deleted || e is MoveFrom || e is MoveTo || e is ParagraphMarkRunPropertiesChange))
                .Select(e => e.CloneNode(true)));
        if (runProperties.HasChildren)
        {
            run.PrependChild(runProperties);
        }
    }

    private void DeleteAllRunsInList(List<Run> list)
    {
        //Note that we must the list of elements that are to be removed in order to avoid modifying the collection that we are iterating over.
        foreach (Run run in list)
        {
            run.Parent.RemoveChild(run);
        }
    }

    /// <summary>
    /// A new range over the same elements (<c>Range.Duplicate</c> in Word).
    /// </summary>
    public RangeWrapper Duplicate()
    {
        if (!_isSpan)
        {
            return new RangeWrapper(_startOpenXmlElement, _endOpenXmlElement);
        }

        RangeWrapper copy = new RangeWrapper(_startOpenXmlElement ?? _root);
        copy._isSpan = true;
        copy._root = _root;
        copy._startAnchor = _startAnchor.Clone();
        copy._endAnchor = _endAnchor.Clone();
        copy._lastStart = _lastStart;
        copy._lastEnd = _lastEnd;
        copy._startOpenXmlElement = null;
        copy._endOpenXmlElement = null;
        SpanRegistry.Register(_root, copy);
        return copy;
    }

    public ParagraphWrapper FirstParagraph
    {
        get { return Paragraphs[0]; }
    }

    public ParagraphWrapper LastParagraph
    {
        get { return Paragraphs[Paragraphs.Count - 1]; }
    }

    #region Private Static Helpers

    private void AppendTableText(StringBuilder sb, Table table)
    {
        foreach (TableRow tableRow in table.Elements<TableRow>())
        {
            AppendTableRowText(sb, tableRow);
        }
    }

    private void AppendTableRowText(StringBuilder sb, TableRow tableRow)
    {
        foreach (TableCell tableCell in tableRow.Elements<TableCell>())
        {
            AppendTableCellText(sb, tableCell);
        }
    }

    private void AppendTableCellText(StringBuilder sb, TableCell tableCell)
    {
        foreach (Paragraph paragraph in tableCell.Elements<Paragraph>())
        {
            AppendParagraphText(sb, paragraph);
        }

        //If the table cell is the last one in the table row then add the end of row marker.
        if (tableCell.Parent.LastChild.Equals(tableCell))
        {
            sb.Append("\r\a");
        }
    }

    private void AppendParagraphText(StringBuilder sb, Paragraph paragraph)
    {
        //Runs can be nested inside other elements (hyperlinks, content controls, smart tags etc.) but runs that belong to a nested paragraph
        //(for example in a text box) are not part of this paragraph.
        foreach (Run run in paragraph.Descendants<Run>().Where(r => r.Ancestors<Paragraph>().First().Equals(paragraph)))
        {
            AppendRunText(sb, run, false);
        }

        //Every paragraph, including an empty one, ends with an end of paragraph marker.
        sb.Append('\r');

        //If the paragraph is the last one in a table cell then add the end of cell marker.
        if (paragraph.Parent is TableCell && paragraph.Parent.LastChild.Equals(paragraph))
        {
            sb.Append('\a');
        }
    }

    private void AppendRunText(StringBuilder sb, Run run, bool appendParagraphMarker)
    {
        foreach (OpenXmlElement element in run.Elements())
        {
            if (element is Text)
            {
                sb.Append(((Text)element).Text);
            }
            else if (element is TabChar)
            {
                sb.Append('\t');
            }
            else if (element is Break)
            {
                //Word represents a page break as a form feed and any other break as a vertical tab.
                Break lineBreak = (Break)element;
                sb.Append(lineBreak.Type != null && lineBreak.Type.HasValue && lineBreak.Type.Value == BreakValues.Page ? '\f' : '\v');
            }
        }

        //If this is the last run with text in the paragraph append an end of paragraph marker. This is only used when a range
        //(bookmark or comment) is being read since a whole paragraph appends its own marker.
        if (appendParagraphMarker && run.IsLastWithText() && !(_startOpenXmlElement is BookmarkStart))
        {
            sb.Append('\r');
        }
    }

    private void AppendBookmarkText(StringBuilder sb, BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd)
    {
        foreach (Run run in bookmarkStart.GetElements<Run>(bookmarkEnd))
        {
            AppendRunText(sb, run, true);
        }
    }

    private void AppendCommentText(StringBuilder sb, CommentRangeStart commentRangeStart, CommentRangeEnd commentRangeEnd)
    {
        foreach (Run run in commentRangeStart.GetElements<Run>(commentRangeEnd))
        {
            AppendRunText(sb, run, true);
        }
    }

    private void AppendCommentText(StringBuilder sb, Comment comment)
    {
        foreach (Paragraph paragraph in comment.GetElements<Paragraph>())
        {
            AppendParagraphText(sb, paragraph);
        }
    }

    #endregion

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as RangeWrapper);
    }

    public bool Equals(RangeWrapper other)
    {
        if (ReferenceEquals(other, null))
        {
            return false;
        }
        if (_isSpan || other._isSpan)
        {
            return _isSpan && other._isSpan && ReferenceEquals(_root, other._root) && _startAnchor.SameAs(other._startAnchor) && _endAnchor.SameAs(other._endAnchor);
        }
        return ReferenceEquals(_startOpenXmlElement, other._startOpenXmlElement) && ReferenceEquals(_endOpenXmlElement, other._endOpenXmlElement);
    }

    public override int GetHashCode()
    {
        if (_isSpan)
        {
            return System.HashCode.Combine(System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_root), _startAnchor.Element == null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_startAnchor.Element), _startAnchor.Offset, _endAnchor.Element == null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_endAnchor.Element), _endAnchor.Offset);
        }
        return System.HashCode.Combine(System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_startOpenXmlElement), System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_endOpenXmlElement));
    }

    public static bool operator ==(RangeWrapper left, RangeWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(RangeWrapper left, RangeWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
