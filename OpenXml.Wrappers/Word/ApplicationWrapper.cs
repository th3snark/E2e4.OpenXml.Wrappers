using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The application, as <c>Application</c> in Word: the open documents, the active one and the name of the user. There is no program behind it;
/// it only keeps the documents together. A document opened or created through <see cref="WordprocessingDocumentWrapper.Open(string)"/> or
/// <see cref="WordprocessingDocumentWrapper.Create(string, bool, bool)"/> belongs to an application of its own, which
/// <see cref="WordprocessingDocumentWrapper.Application"/> gives.
/// </summary>
public partial class ApplicationWrapper : IDisposable
{
    private readonly List<WordprocessingDocumentWrapper> _documents = new List<WordprocessingDocumentWrapper>();
    private WordprocessingDocumentWrapper _active;
    private int _lastNumber;
    private string _userName;
    private string _userInitials;
    private string _userAddress = string.Empty;
    private bool _quit;

    /// <summary>Starts an application with no open documents. The user is the user of the system.</summary>
    public ApplicationWrapper()
    {
        _userName = string.IsNullOrEmpty(Environment.UserName) ? "Unknown" : Environment.UserName;
        _userInitials = CommentParts.GetInitials(_userName);
    }

    /// <summary>The name of the application, as <c>Application.Name</c> in Word.</summary>
    public string Name
    {
        get { return "Microsoft Word"; }
    }

    /// <summary>
    /// The name of the user, as <c>Application.UserName</c> in Word. It is the author of the comments added to the documents of the
    /// application. It starts as the name of the user of the system.
    /// </summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public string UserName
    {
        get { return _userName; }
        set
        {
            Guard.ThrowIfNull(value);
            //The initials do not follow the name, as in Word.
            _userName = value;
        }
    }

    /// <summary>The initials of the user, as <c>Application.UserInitials</c> in Word. They are the initials of the comments.</summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public string UserInitials
    {
        get { return _userInitials; }
        set
        {
            Guard.ThrowIfNull(value);
            _userInitials = value;
        }
    }

    /// <summary>
    /// The version of Word whose object model the library follows, as <c>Application.Version</c> in Word (it is "16.0" for Word 2016 to 365).
    /// </summary>
    public string Version
    {
        get { return "16.0"; }
    }

    /// <summary>The address of the user, as <c>Application.UserAddress</c> in Word; empty until it is set.</summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public string UserAddress
    {
        get { return _userAddress; }
        set
        {
            Guard.ThrowIfNull(value);
            _userAddress = value;
        }
    }

    /// <summary>Centimetres in points, as <c>Application.CentimetersToPoints</c> in Word.</summary>
    public double CentimetersToPoints(double centimeters)
    {
        return centimeters * 72 / 2.54;
    }

    /// <summary>Inches in points, as <c>Application.InchesToPoints</c> in Word.</summary>
    public double InchesToPoints(double inches)
    {
        return inches * 72;
    }

    /// <summary>Lines in points (a line is 12 points), as <c>Application.LinesToPoints</c> in Word.</summary>
    public double LinesToPoints(double lines)
    {
        return lines * 12;
    }

    /// <summary>Millimetres in points, as <c>Application.MillimetersToPoints</c> in Word.</summary>
    public double MillimetersToPoints(double millimeters)
    {
        return millimeters * 72 / 25.4;
    }

    /// <summary>Picas in points (a pica is 12 points), as <c>Application.PicasToPoints</c> in Word.</summary>
    public double PicasToPoints(double picas)
    {
        return picas * 12;
    }

    /// <summary>Points in centimetres, as <c>Application.PointsToCentimeters</c> in Word.</summary>
    public double PointsToCentimeters(double points)
    {
        return points * 2.54 / 72;
    }

    /// <summary>Points in inches, as <c>Application.PointsToInches</c> in Word.</summary>
    public double PointsToInches(double points)
    {
        return points / 72;
    }

    /// <summary>Points in lines (a line is 12 points), as <c>Application.PointsToLines</c> in Word.</summary>
    public double PointsToLines(double points)
    {
        return points / 12;
    }

    /// <summary>Points in millimetres, as <c>Application.PointsToMillimeters</c> in Word.</summary>
    public double PointsToMillimeters(double points)
    {
        return points * 25.4 / 72;
    }

    /// <summary>Points in picas (a pica is 12 points), as <c>Application.PointsToPicas</c> in Word.</summary>
    public double PointsToPicas(double points)
    {
        return points / 12;
    }

    /// <summary>The open documents, as <c>Application.Documents</c> in Word.</summary>
    public DocumentWrapperCollection Documents
    {
        get { return new DocumentWrapperCollection(this); }
    }

    /// <summary>The number of documents, as <c>Application.Documents.Count</c> in Word.</summary>
    internal int Count
    {
        get { return _documents.Count; }
    }

    /// <summary>
    /// The document that was last created, opened or activated and has not been closed since, as <c>Application.ActiveDocument</c> in Word.
    /// </summary>
    /// <exception cref="InvalidOperationException">No document is open.</exception>
    public WordprocessingDocumentWrapper ActiveDocument
    {
        get { return _active ?? throw new InvalidOperationException("No document is open."); }
    }

    /// <summary>The selection of the active document, as <c>Application.Selection</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">No document is open.</exception>
    public SelectionWrapper Selection
    {
        get { return ActiveDocument.Selection; }
    }

    /// <summary>
    /// Closes every document, as <c>Application.Quit</c> in Word. The documents are saved when they were opened so that they can be changed,
    /// as when each is closed.
    /// </summary>
    public void Quit()
    {
        Dispose();
    }

    /// <summary>
    /// Closes every document, saving the changes or not, as <c>Application.Quit</c> in Word.
    /// </summary>
    /// <exception cref="NotSupportedException"><see cref="SaveOptions.PromptToSaveChanges"/>, because there is no user to ask.</exception>
    public void Quit(SaveOptions saveChanges)
    {
        if (!EnumCompat.IsDefined(saveChanges))
        {
            throw new ArgumentOutOfRangeException(nameof(saveChanges));
        }
        if (saveChanges == SaveOptions.PromptToSaveChanges)
        {
            throw new NotSupportedException("There is no user to ask whether to save the changes.");
        }
        foreach (WordprocessingDocumentWrapper document in _documents.ToList())
        {
            document.Close(saveChanges);
        }
        _quit = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (!_quit)
        {
            _quit = true;
            foreach (WordprocessingDocumentWrapper document in _documents.ToList())
            {
                document.Close();
            }
        }
        GC.SuppressFinalize(this);
    }

    internal IList<WordprocessingDocumentWrapper> DocumentList
    {
        get { return _documents; }
    }

    //The name Word gives a new document: Document1, Document2 and so on, a number being used once.
    internal string NewDocumentName()
    {
        _lastNumber++;
        return string.Format(CultureInfo.InvariantCulture, "Document{0}", _lastNumber);
    }

    internal void Register(WordprocessingDocumentWrapper document)
    {
        //The newest document is the first, as in Word.
        _documents.Insert(0, document);
        _active = document;
    }

    internal void Unregister(WordprocessingDocumentWrapper document)
    {
        _documents.Remove(document);
        if (ReferenceEquals(_active, document))
        {
            _active = _documents.FirstOrDefault();
        }
    }

    internal void Activate(WordprocessingDocumentWrapper document)
    {
        _active = document;
    }
}

/// <summary>
/// The open documents of an application, as <c>Documents</c> in Word. The newest document is the first.
/// </summary>
public class DocumentWrapperCollection : IEnumerable<WordprocessingDocumentWrapper>
{
    private readonly ApplicationWrapper _application;

    internal DocumentWrapperCollection(ApplicationWrapper application)
    {
        _application = application;
    }

    /// <summary>The application the documents are in.</summary>
    public ApplicationWrapper Application
    {
        get { return _application; }
    }

    /// <summary>The number of open documents.</summary>
    public int Count
    {
        get { return _application.Count; }
    }

    /// <summary>The document at the index (0-based), as <c>Documents.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no document at the index.</exception>
    public WordprocessingDocumentWrapper this[int index]
    {
        get
        {
            if (index < 0 || index >= _application.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return _application.DocumentList[index];
        }
    }

    /// <summary>The document with the name (the name of its file, or Document1 and so on), as <c>Documents.Item</c> in Word.</summary>
    /// <exception cref="ArgumentNullException">The name is null.</exception>
    /// <exception cref="KeyNotFoundException">No open document has the name.</exception>
    public WordprocessingDocumentWrapper this[string name]
    {
        get
        {
            Guard.ThrowIfNull(name);
            return _application.DocumentList.FirstOrDefault(d => string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase))
                ?? throw new KeyNotFoundException("No open document is called " + name + ".");
        }
    }

    /// <summary>
    /// Creates a new blank document that is not saved anywhere yet and makes it the active document, as <c>Documents.Add</c> in Word. It is
    /// called Document1, Document2 and so on; save it with <see cref="WordprocessingDocumentWrapper.SaveAs(string)"/>.
    /// </summary>
    public WordprocessingDocumentWrapper Add()
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(new MemoryStream(), true, false);
        document.Join(_application, _application.NewDocumentName());
        return document;
    }

    /// <summary>
    /// Opens a document from a file and makes it the active document, as <c>Documents.Open</c> in Word. A file that is open already is not
    /// opened again; the open document is made the active one.
    /// </summary>
    /// <param name="fileName">The path of the file.</param>
    /// <param name="readOnly">True to open it so that it cannot be changed.</param>
    public WordprocessingDocumentWrapper Open(string fileName, bool readOnly = false)
    {
        Guard.ThrowIfNull(fileName);
        string full = Path.GetFullPath(fileName);
        WordprocessingDocumentWrapper open = _application.DocumentList.FirstOrDefault(d => d.Path.Length > 0
            && string.Equals(d.FullName, full, StringComparison.OrdinalIgnoreCase));
        if (open != null)
        {
            _application.Activate(open);
            return open;
        }

        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(fileName, !readOnly);
        document.Join(_application, null);
        return document;
    }

    /// <summary>Saves every open document that can be saved to its file, as <c>Documents.Save</c> in Word.</summary>
    public void Save()
    {
        foreach (WordprocessingDocumentWrapper document in _application.DocumentList.Where(d => !d.ReadOnly))
        {
            document.Save();
        }
    }

    /// <summary>Closes every open document, saving the changes or not, as <c>Documents.Close</c> in Word.</summary>
    /// <exception cref="NotSupportedException"><see cref="SaveOptions.PromptToSaveChanges"/>, because there is no user to ask.</exception>
    public void Close(SaveOptions saveChanges = SaveOptions.SaveChanges)
    {
        if (!EnumCompat.IsDefined(saveChanges))
        {
            throw new ArgumentOutOfRangeException(nameof(saveChanges));
        }
        if (saveChanges == SaveOptions.PromptToSaveChanges)
        {
            throw new NotSupportedException("There is no user to ask whether to save the changes.");
        }
        foreach (WordprocessingDocumentWrapper document in _application.DocumentList.ToList())
        {
            document.Close(saveChanges);
        }
    }

    /// <inheritdoc/>
    public IEnumerator<WordprocessingDocumentWrapper> GetEnumerator()
    {
        return _application.DocumentList.ToList().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
