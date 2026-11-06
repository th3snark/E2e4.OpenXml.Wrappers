namespace OpenXml.Wrappers.Word;

public partial class WordprocessingDocumentWrapper
{
    private SelectionWrapper _selection;

    /// <summary>
    /// The selection of the document, as <c>Selection</c> (of the active window) in Word: an insertion point at the start of the main text
    /// until something selects another place. There is one for each document.
    /// </summary>
    public SelectionWrapper Selection
    {
        get { return _selection ?? (_selection = new SelectionWrapper(this)); }
    }
}
