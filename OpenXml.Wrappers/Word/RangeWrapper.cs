using System;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class RangeWrapper
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

    public OpenXmlElement Start
    {
        get { return _startOpenXmlElement; }
    }

    public OpenXmlElement End
    {
        get { return _endOpenXmlElement; }
    }

    public BookmarkWrapperCollection Bookmarks
    {
        get { return new BookmarkWrapperCollection(this); }
    }

    public TableWrapperCollection Tables
    {
        get { return new TableWrapperCollection(this); }
    }

    public TableCellWrapperCollection Cells
    {
        get { return new TableCellWrapperCollection(this); }
    }

    public TableRowWrapperCollection Rows
    {
        get { return new TableRowWrapperCollection(this); }
    }

    public ParagraphWrapperCollection Paragraphs
    {
        get { return new ParagraphWrapperCollection(this); }
    }

    public CommentWrapperCollection Comments
    {
        get { return new CommentWrapperCollection(this); }
    }

    public string Text
    {
        get
        {
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
            else
            {
                throw new InvalidOperationException(String.Format(CultureInfo.InvariantCulture, "Getting the Text property is not currently supported for the type {0}", _startOpenXmlElement.GetType().Name));
            }
            return sb.ToString();
        }

        set
        {
            string text = value == null ? String.Empty : value;

            if (_startOpenXmlElement is TableCell)
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
                        run.AppendChild<Text>(new Text(sb.ToString()));
                    }
                    if (paragraphMarkRunProperties != null)
                    {
                        run.PrependChild<ParagraphMarkRunProperties>((ParagraphMarkRunProperties)paragraphMarkRunProperties.Clone());
                    }
                    runList.Insert(0, run);
                    sb.Length = 0;
                    break;

                case '\n':
                case '\v':
                    run = new Run(new Break());
                    if (sb.Length > 0)
                    {
                        run.AppendChild<Text>(new Text(sb.ToString()));
                    }
                    if (paragraphMarkRunProperties != null)
                    {
                        run.PrependChild<ParagraphMarkRunProperties>((ParagraphMarkRunProperties)paragraphMarkRunProperties.Clone());
                    }
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
            run = new Run(new Text(sb.ToString()));
            if (paragraphMarkRunProperties != null)
            {
                run.PrependChild<ParagraphMarkRunProperties>((ParagraphMarkRunProperties)paragraphMarkRunProperties.Clone());
            }

            runList.Insert(0, run);
        }
        return runList;
    }

    private void DeleteAllRunsInList(List<Run> list)
    {
        //Note that we must the list of elements that are to be removed in order to avoid modifying the collection that we are iterating over.
        foreach (Run run in list)
        {
            run.Parent.RemoveChild(run);
        }
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
        foreach (Run run in paragraph.Elements<Run>())
        {
            AppendRunText(sb, run);
        }

        //If the paragraph is the last one in a table cell then add the end of cell marker.
        if (paragraph.Parent is TableCell && paragraph.Parent.LastChild.Equals(paragraph))
        {
            sb.Append('\a');
        }
    }

    private void AppendRunText(StringBuilder sb, Run run)
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
        }

        //If this is the last run with text in the paragraph append an end of paragraph marker.
        if (run.IsLastWithText() && !(_startOpenXmlElement is BookmarkStart))
        {
            sb.Append('\r');
        }
    }

    private void AppendBookmarkText(StringBuilder sb, BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd)
    {
        foreach (Run run in bookmarkStart.GetElements<Run>(bookmarkEnd))
        {
            AppendRunText(sb, run);
        }
    }

    private void AppendCommentText(StringBuilder sb, CommentRangeStart commentRangeStart, CommentRangeEnd commentRangeEnd)
    {
        foreach (Run run in commentRangeStart.GetElements<Run>(commentRangeEnd))
        {
            AppendRunText(sb, run);
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
}
