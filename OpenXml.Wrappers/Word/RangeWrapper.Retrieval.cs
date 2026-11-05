using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The document a range is in and what its text includes.
/// </summary>
public partial class RangeWrapper
{
    private TextRetrievalMode _retrieval;

    /// <summary>
    /// What <see cref="Text"/> includes, as <c>Range.TextRetrievalMode</c> in Word: field codes and hidden text. The settings belong to this
    /// range; <see cref="Duplicate"/> copies them.
    /// </summary>
    public TextRetrievalMode TextRetrievalMode
    {
        get { return _retrieval ?? (_retrieval = new TextRetrievalMode()); }
    }

    /// <summary>
    /// The document the range is in, as <c>Range.Document</c> in Word: the wrapper the document was opened or created with. Null if the range
    /// is over elements that are not in a document, or the document was not opened with a <see cref="WordprocessingDocumentWrapper"/>.
    /// </summary>
    public WordprocessingDocumentWrapper Document
    {
        get
        {
            OpenXmlElement root = StoryMap.FindRoot(AnchorElement);
            OpenXmlPart part = (root?.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart;
            return WordprocessingDocumentWrapper.For(part?.OpenXmlPackage as WordprocessingDocument);
        }
    }

    /// <summary>
    /// The object the range belongs to, as <c>Range.Parent</c> in Word: the document (a <see cref="WordprocessingDocumentWrapper"/>), or
    /// null where <see cref="Document"/> is.
    /// </summary>
    public object Parent
    {
        get { return Document; }
    }
}
