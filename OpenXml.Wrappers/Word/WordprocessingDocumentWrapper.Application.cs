namespace OpenXml.Wrappers.Word;

public partial class WordprocessingDocumentWrapper
{
    private ApplicationWrapper _application;
    private string _untitledName;

    //The application, if the document has one yet. Documents get theirs when it is asked for or when they are added through Documents.
    internal ApplicationWrapper ApplicationIfAny
    {
        get { return _application; }
    }

    /// <summary>
    /// The application the document is open in, as <c>Document.Application</c> in Word. A document that was not added or opened through
    /// <see cref="ApplicationWrapper.Documents"/> has an application of its own, made when this is first read, in which it is the only
    /// document.
    /// </summary>
    public ApplicationWrapper Application
    {
        get
        {
            if (_application == null)
            {
                ApplicationWrapper application = new ApplicationWrapper();
                Join(application, null);
            }
            return _application;
        }
    }

    internal void Join(ApplicationWrapper application, string untitledName)
    {
        _application = application;
        //A document without a file takes the next name of the application, unless it has been called something already.
        _untitledName = untitledName ?? (DocumentFiles.TryGetPath(_wordprocessingDocument, out _) ? null : _untitledName);
        application.Register(this);
    }

    /// <summary>Makes the document the active document of its application, as <c>Document.Activate</c> in Word.</summary>
    public void Activate()
    {
        Application.Activate(this);
    }
}
