using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Whether to save the changes when a document is closed, as <c>WdSaveOptions</c> in Word (with the same values).
/// </summary>
public enum SaveOptions
{
    DoNotSaveChanges = 0,
    SaveChanges = -1,
    PromptToSaveChanges = -2
}

/// <summary>
/// The format a document is saved in, as <c>WdSaveFormat</c> in Word (with the same values). Only the Open XML formats are supported.
/// </summary>
public enum SaveFormat
{
    /// <summary>A document (.docx).</summary>
    XmlDocument = 12,

    /// <summary>A document that can hold macros (.docm).</summary>
    XmlDocumentMacroEnabled = 13,

    /// <summary>A template (.dotx).</summary>
    XmlTemplate = 14,

    /// <summary>A template that can hold macros (.dotm).</summary>
    XmlTemplateMacroEnabled = 15
}

/// <summary>
/// What a document is protected against, as <c>WdProtectionType</c> in Word (with the same values).
/// </summary>
public enum ProtectionType
{
    NoProtection = -1,
    AllowOnlyRevisions = 0,
    AllowOnlyComments = 1,
    AllowOnlyFormFields = 2,
    AllowOnlyReading = 3
}

/// <summary>
/// A statistic of a document, as <c>WdStatistic</c> in Word (with the same values).
/// </summary>
public enum Statistic
{
    Words = 0,

    /// <summary>The lines. Not supported: it depends on the layout of the pages.</summary>
    Lines = 1,

    /// <summary>The pages. Not supported: it depends on the layout of the pages.</summary>
    Pages = 2,

    /// <summary>The characters not counting spaces.</summary>
    Characters = 3,
    Paragraphs = 4,
    CharactersWithSpaces = 5
}

/// <summary>
/// Where the path of a document is kept: the library knows it, the package does not.
/// </summary>
internal static class DocumentFiles
{
    private static readonly ConditionalWeakTable<WordprocessingDocument, string> Paths = new ConditionalWeakTable<WordprocessingDocument, string>();

    public static void Remember(WordprocessingDocument document, string path)
    {
        Paths.Remove(document);
        if (path != null)
        {
            Paths.Add(document, path);
        }
    }

    public static bool TryGetPath(WordprocessingDocument document, out string path)
    {
        return Paths.TryGetValue(document, out path);
    }
}

public partial class WordprocessingDocumentWrapper
{
    /// <summary>
    /// The name of the file of the document, as <c>Document.Name</c> in Word. A document that was not opened from or saved to a file is
    /// called "Document1".
    /// </summary>
    public string Name
    {
        get { return DocumentFiles.TryGetPath(_wordprocessingDocument, out string path) ? System.IO.Path.GetFileName(path) : _untitledName ?? "Document1"; }
    }

    /// <summary>The full path of the file of the document, as <c>Document.FullName</c> in Word.</summary>
    public string FullName
    {
        get { return DocumentFiles.TryGetPath(_wordprocessingDocument, out string path) ? System.IO.Path.GetFullPath(path) : Name; }
    }

    /// <summary>The folder of the file of the document, as <c>Document.Path</c> in Word; empty for a document without a file.</summary>
    public string Path
    {
        get { return DocumentFiles.TryGetPath(_wordprocessingDocument, out string path) ? System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path)) : string.Empty; }
    }

    /// <summary>True if the document was opened so that it cannot be changed, as <c>Document.ReadOnly</c> in Word.</summary>
    public bool ReadOnly
    {
        get { return _wordprocessingDocument.FileOpenAccess == FileAccess.Read; }
    }

    /// <summary>
    /// Closes the document, saving the changes or not, as <c>Document.Close</c> in Word.
    /// </summary>
    /// <exception cref="NotSupportedException"><see cref="SaveOptions.PromptToSaveChanges"/>, because there is no user to ask.</exception>
    public void Close(SaveOptions saveChanges)
    {
        if (!EnumCompat.IsDefined(saveChanges))
        {
            throw new ArgumentOutOfRangeException(nameof(saveChanges));
        }
        if (saveChanges == SaveOptions.PromptToSaveChanges)
        {
            throw new NotSupportedException("There is no user to ask whether to save the changes.");
        }

        if (saveChanges == SaveOptions.SaveChanges)
        {
            SaveChanges();
        }
        else if (_wordprocessingDocument.AutoSave && _wordprocessingDocument.FileOpenAccess != FileAccess.Read)
        {
            throw new NotSupportedException("The package saves itself when it is closed because it was not opened by this library; use Open or Create of the wrapper.");
        }
        else
        {
            _skipSave = true;
        }
        Dispose();
    }

    /// <summary>
    /// Saves a copy of the document to a new file in the format, as <c>Document.SaveAs2</c> in Word, and returns a wrapper over the copy.
    /// </summary>
    public WordprocessingDocumentWrapper SaveAs(string path, SaveFormat fileFormat)
    {
        if (!EnumCompat.IsDefined(fileFormat))
        {
            throw new ArgumentOutOfRangeException(nameof(fileFormat));
        }

        WordprocessingDocumentWrapper copy = SaveAs(path);
        WordprocessingDocumentType type = fileFormat == SaveFormat.XmlDocumentMacroEnabled ? WordprocessingDocumentType.MacroEnabledDocument
            : fileFormat == SaveFormat.XmlTemplate ? WordprocessingDocumentType.Template
            : fileFormat == SaveFormat.XmlTemplateMacroEnabled ? WordprocessingDocumentType.MacroEnabledTemplate : WordprocessingDocumentType.Document;
        if (copy.WordprocessingDocument.DocumentType != type)
        {
            copy.WordprocessingDocument.ChangeDocumentType(type);
        }
        return copy;
    }

    /// <summary>
    /// The page setup that applies to every section of the document, as <c>Document.PageSetup</c> in Word. Reading gives the value of the
    /// first section and setting sets it in all of them.
    /// </summary>
    public DocumentPageSetupWrapper PageSetup
    {
        get { return new DocumentPageSetupWrapper(Sections); }
    }

    /// <summary>The characters of the document, as <c>Document.Characters</c> in Word.</summary>
    public TextUnitCollection Characters
    {
        get { return Content.Characters; }
    }

    /// <summary>The words of the document, as <c>Document.Words</c> in Word.</summary>
    public TextUnitCollection Words
    {
        get { return Content.Words; }
    }

    /// <summary>The sentences of the document, as <c>Document.Sentences</c> in Word.</summary>
    public TextUnitCollection Sentences
    {
        get { return Content.Sentences; }
    }

    /// <summary>
    /// The content of the document as a package of Office Open XML in one piece, as <c>Document.WordOpenXML</c> in Word. See
    /// <see cref="RangeWrapper.WordOpenXML"/>.
    /// </summary>
    public string WordOpenXML
    {
        get { return Content.WordOpenXML; }
    }

    /// <summary>
    /// Counts something in the main text of the document, as <c>Document.ComputeStatistics</c> in Word.
    /// </summary>
    /// <exception cref="NotSupportedException">The statistic depends on the layout of the pages (<see cref="Statistic.Lines"/> and <see cref="Statistic.Pages"/>).</exception>
    public int ComputeStatistics(Statistic statistic)
    {
        if (!EnumCompat.IsDefined(statistic))
        {
            throw new ArgumentOutOfRangeException(nameof(statistic));
        }
        if (statistic == Statistic.Lines || statistic == Statistic.Pages)
        {
            throw new NotSupportedException("Lines and pages depend on the layout of the pages, which the library does not work out.");
        }

        string text = Content.Text ?? string.Empty;
        if (statistic == Statistic.Words)
        {
            return text.Split(new[] { ' ', '\t', '\r', '\n', '\a', '\v', '\f', ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }
        if (statistic == Statistic.Paragraphs)
        {
            return Paragraphs.Count;
        }
        if (statistic == Statistic.Characters)
        {
            return text.Count(c => !char.IsWhiteSpace(c) && !char.IsControl(c));
        }
        return text.Count(c => !char.IsControl(c) || c == '\t');
    }

    private Settings GetSettings(bool create)
    {
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart;
        DocumentSettingsPart part = main.DocumentSettingsPart;
        if (part == null && create)
        {
            part = main.AddNewPart<DocumentSettingsPart>();
        }
        if (part != null && part.Settings == null && create)
        {
            part.Settings = new Settings();
        }
        return part?.Settings;
    }

    /// <summary>
    /// True if changes are tracked as revisions, as <c>Document.TrackRevisions</c> in Word. Reading it works for any document. Turning it on is
    /// not supported yet, because edits made through this library are not recorded as revisions (tracked insertions and deletions), so a
    /// document that says it tracks changes would be edited without tracking: the setter throws instead. Turning it off is allowed.
    /// </summary>
    /// <exception cref="NotSupportedException">Setting it to true when it is off.</exception>
    public bool TrackRevisions
    {
        get
        {
            TrackRevisions setting = GetSettings(false)?.GetFirstChild<TrackRevisions>();
            return setting != null && (setting.Val == null || !setting.Val.HasValue || setting.Val.Value);
        }
        set
        {
            if (!value)
            {
                GetSettings(false)?.RemoveAllChildren<TrackRevisions>();
            }
            else if (!TrackRevisions)
            {
                throw new NotSupportedException("Tracking changes is not supported yet: edits made through the library are not recorded as revisions.");
            }
        }
    }

    /// <summary>
    /// The name of the default table style that the document says it has (<c>w:defaultTableStyle</c>), as <c>Document.DefaultTableStyle</c>
    /// in Word, or null when none is set (Word gives a style without a name then). It is a setting of the document: it does not change the
    /// style of the tables already there, nor of those <c>Tables.Add</c> makes, which Word also leaves alone.
    /// </summary>
    public string DefaultTableStyle
    {
        get
        {
            string id = GetSettings(false)?.GetFirstChild<DefaultTableStyle>()?.Val?.Value;
            return id == null ? null : StyleHelper.GetStyleName(_wordprocessingDocument.MainDocumentPart.Document, id);
        }
    }

    /// <summary>The default table style as an object (see <see cref="DefaultTableStyle"/>); null when none is set.</summary>
    public StyleWrapper DefaultTableStyleObject
    {
        get { return StyleHelper.GetStyleObject(_wordprocessingDocument.MainDocumentPart.Document, DefaultTableStyle); }
    }

    /// <summary>
    /// Sets the default table style of the document, as <c>Document.SetDefaultTableStyle</c> in Word. Word writes it even for "Table Normal".
    /// </summary>
    /// <param name="style">The name of a table style.</param>
    /// <param name="setInTemplate">Whether to set it in the template as well; there is no template here, so it has no effect.</param>
    /// <exception cref="ArgumentException">There is no table style with the name.</exception>
    public void SetDefaultTableStyle(string style, bool setInTemplate)
    {
        Guard.ThrowIfNull(style, nameof(style));

        string id = StyleHelper.FindStyleId(_wordprocessingDocument.MainDocumentPart.Document, style, out StyleValues type);
        if (type != StyleValues.Table)
        {
            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "The style '{0}' is not a table style.", style), nameof(style));
        }

        Settings settings = GetSettings(true);
        settings.RemoveAllChildren<DefaultTableStyle>();
        settings.AddChild(new DefaultTableStyle() { Val = id });
    }

    /// <summary>Sets the default table style of the document to a style of it; see <see cref="SetDefaultTableStyle(string, bool)"/>.</summary>
    public void SetDefaultTableStyle(StyleWrapper style, bool setInTemplate)
    {
        Guard.ThrowIfNull(style, nameof(style));
        SetDefaultTableStyle(style.Name, setInTemplate);
    }

    /// <summary>
    /// The compatibility mode of the document: 15 for Word 2013 and later, 14 for Word 2010, 12 for Word 2007 (<c>Document.CompatibilityMode</c>
    /// in Word). Documents without a setting are in mode 12.
    /// </summary>
    public int CompatibilityMode
    {
        get
        {
            CompatibilitySetting setting = GetSettings(false)?.GetFirstChild<Compatibility>()?.Elements<CompatibilitySetting>()
                .FirstOrDefault(s => s.Name != null && s.Name.HasValue && s.Name.Value == CompatSettingNameValues.CompatibilityMode);
            return setting?.Val != null && int.TryParse(setting.Val.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int mode) ? mode : 12;
        }
        set
        {
            if (value != 11 && value != 12 && value != 14 && value != 15)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The mode must be 11, 12, 14 or 15.");
            }

            Settings settings = GetSettings(true);
            Compatibility compatibility = settings.GetFirstChild<Compatibility>();
            if (compatibility == null)
            {
                compatibility = new Compatibility();
                settings.AddChild(compatibility);
            }
            CompatibilitySetting setting = compatibility.Elements<CompatibilitySetting>().FirstOrDefault(s => s.Name != null && s.Name.HasValue && s.Name.Value == CompatSettingNameValues.CompatibilityMode);
            if (setting == null)
            {
                setting = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word" };
                compatibility.AppendChild(setting);
            }
            setting.Val = value.ToString(CultureInfo.InvariantCulture);
        }
    }

    #region Protection

    /// <summary>What the document is protected against, as <c>Document.ProtectionType</c> in Word.</summary>
    public ProtectionType ProtectionType
    {
        get
        {
            DocumentProtection protection = GetSettings(false)?.GetFirstChild<DocumentProtection>();
            if (protection == null || protection.Enforcement == null || !protection.Enforcement.HasValue || !protection.Enforcement.Value)
            {
                return ProtectionType.NoProtection;
            }
            DocumentProtectionValues? edit = protection.Edit?.Value;
            if (edit == DocumentProtectionValues.TrackedChanges)
            {
                return ProtectionType.AllowOnlyRevisions;
            }
            if (edit == DocumentProtectionValues.Comments)
            {
                return ProtectionType.AllowOnlyComments;
            }
            return edit == DocumentProtectionValues.Forms ? ProtectionType.AllowOnlyFormFields : ProtectionType.AllowOnlyReading;
        }
    }

    /// <summary>
    /// Protects the document against changes, as <c>Document.Protect</c> in Word. With a password the document can only be unprotected with
    /// the same password.
    /// </summary>
    /// <exception cref="InvalidOperationException">The document is protected already.</exception>
    public void Protect(ProtectionType type, string password = null)
    {
        if (type == ProtectionType.NoProtection || !EnumCompat.IsDefined(type))
        {
            throw new ArgumentOutOfRangeException(nameof(type));
        }
        if (ProtectionType != ProtectionType.NoProtection)
        {
            throw new InvalidOperationException("The document is protected already.");
        }

        Settings settings = GetSettings(true);
        settings.RemoveAllChildren<DocumentProtection>();
        DocumentProtection protection = new DocumentProtection()
        {
            Edit = type == ProtectionType.AllowOnlyRevisions ? DocumentProtectionValues.TrackedChanges
                : type == ProtectionType.AllowOnlyComments ? DocumentProtectionValues.Comments
                : type == ProtectionType.AllowOnlyFormFields ? DocumentProtectionValues.Forms : DocumentProtectionValues.ReadOnly,
            Enforcement = OpenXmlExtensions.OnValue()
        };
        if (!string.IsNullOrEmpty(password))
        {
            byte[] salt = new byte[16];
            using (RandomNumberGenerator random = RandomNumberGenerator.Create())
            {
                random.GetBytes(salt);
            }
            protection.CryptographicProviderType = CryptProviderValues.RsaAdvancedEncryptionStandard;
            protection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
            protection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
            protection.CryptographicAlgorithmSid = 14;
            protection.CryptographicSpinCount = 100000;
            protection.Hash = System.Convert.ToBase64String(HashPassword(password, salt, 100000));
            protection.Salt = System.Convert.ToBase64String(salt);
        }
        settings.AddChild(protection);
    }

    /// <summary>
    /// Removes the protection of the document, as <c>Document.Unprotect</c> in Word.
    /// </summary>
    /// <exception cref="UnauthorizedAccessException">The password is not the one the document was protected with.</exception>
    public void Unprotect(string password = null)
    {
        Settings settings = GetSettings(false);
        DocumentProtection protection = settings?.GetFirstChild<DocumentProtection>();
        if (protection == null)
        {
            return;
        }

        if (protection.Hash != null && protection.Salt != null)
        {
            int spinCount = (int)(protection.CryptographicSpinCount?.Value ?? 100000);
            byte[] expected = System.Convert.FromBase64String(protection.Hash.Value);
            byte[] actual = HashPassword(password ?? string.Empty, System.Convert.FromBase64String(protection.Salt.Value), spinCount);
            if (!SameBytes(expected, actual))
            {
                throw new UnauthorizedAccessException("The password is not correct.");
            }
        }
        else if (protection.HashValue != null && protection.SaltValue != null)
        {
            //As version 1.0 of this library wrote it, which Word could not check.
            int spin = (int)(protection.CryptographicSpinCount?.Value ?? 100000);
            byte[] expected = System.Convert.FromBase64String(protection.HashValue.Value);
            byte[] actual = HashPasswordVersion1(password ?? string.Empty, System.Convert.FromBase64String(protection.SaltValue.Value), spin);
            if (!SameBytes(expected, actual))
            {
                throw new UnauthorizedAccessException("The password is not correct.");
            }
        }
        protection.Remove();
    }

    //Compared by hand: with C# 14 the Linq call on arrays binds to a Span method that net48 would need System.Memory for.
    private static bool SameBytes(byte[] a, byte[] b)
    {
        if (a.Length != b.Length)
        {
            return false;
        }
        int difference = 0;
        for (int i = 0; i < a.Length; i++)
        {
            difference |= a[i] ^ b[i];
        }
        return difference == 0;
    }

#pragma warning disable CA1850 // SHA512.HashData does not exist for netstandard2.0.
    //The way Word 2013 and later hashes a password (SHA-512, a salt, and many rounds).
    //As Word does it, checked against documents Word has protected: the old 16 bit key of the password, as the hexadecimal digits of its bytes in
    //UTF-16, after the salt, and then the hash of that with the number of each round.
    private static byte[] HashPassword(string password, byte[] salt, int spinCount)
    {
        byte[] key = BitConverter.GetBytes(GetLegacyPasswordKey(password));
        string digits = string.Concat(key.Select(b => b.ToString("X2", CultureInfo.InvariantCulture)));
        using (SHA512 sha = SHA512.Create())
        {
            byte[] hash = sha.ComputeHash(salt.Concat(Encoding.Unicode.GetBytes(digits)).ToArray());
            for (int i = 0; i < spinCount; i++)
            {
                hash = sha.ComputeHash(hash.Concat(BitConverter.GetBytes(i)).ToArray());
            }
            return hash;
        }
    }

    private static byte[] HashPasswordVersion1(string password, byte[] salt, int spinCount)
    {
        using (SHA512 sha = SHA512.Create())
        {
            byte[] hash = sha.ComputeHash(salt.Concat(Encoding.Unicode.GetBytes(password)).ToArray());
            for (int i = 0; i < spinCount; i++)
            {
                hash = sha.ComputeHash(hash.Concat(BitConverter.GetBytes(i)).ToArray());
            }
            return hash;
        }
    }

    private static readonly ushort[] PasswordInitialCode = { 0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3 };

    private static readonly ushort[][] PasswordMatrix =
    {
        new ushort[] { 0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09 }, new ushort[] { 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF },
        new ushort[] { 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0 }, new ushort[] { 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40 },
        new ushort[] { 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5 }, new ushort[] { 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A },
        new ushort[] { 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9 }, new ushort[] { 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0 },
        new ushort[] { 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC }, new ushort[] { 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10 },
        new ushort[] { 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168 }, new ushort[] { 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C },
        new ushort[] { 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD }, new ushort[] { 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC },
        new ushort[] { 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4 },
    };

    //The 32 bit key that Word has made of a password since Word 97: up to 15 characters, each as one byte.
    private static uint GetLegacyPasswordKey(string password)
    {
        string text = password.Length > 15 ? password.Substring(0, 15) : password;
        if (text.Length == 0)
        {
            return 0;
        }
        byte[] bytes = text.Select(c => (byte)((c & 0xFF) != 0 ? c & 0xFF : c >> 8)).ToArray();

        uint high = PasswordInitialCode[bytes.Length - 1];
        for (int i = 0; i < bytes.Length; i++)
        {
            ushort[] row = PasswordMatrix[15 - bytes.Length + i];
            for (int bit = 0; bit < 7; bit++)
            {
                if ((bytes[i] & (1 << bit)) != 0)
                {
                    high ^= row[bit];
                }
            }
        }

        int low = 0;
        for (int i = bytes.Length - 1; i >= 0; i--)
        {
            low = (((low >> 14) & 1) | ((low << 1) & 0x7FFF)) ^ bytes[i];
        }
        low = (((low >> 14) & 1) | ((low << 1) & 0x7FFF)) ^ bytes.Length ^ 0xCE4B;
        return (high << 16) | (uint)(low & 0xFFFF);
    }
#pragma warning restore CA1850

    #endregion

    /// <summary>The variables stored in the document, as <c>Document.Variables</c> in Word.</summary>
    public VariableWrapperCollection Variables
    {
        get { return new VariableWrapperCollection(this); }
    }

    internal Settings GetSettingsPart(bool create)
    {
        return GetSettings(create);
    }
}

/// <summary>
/// A variable stored in a document, as <c>Variable</c> in Word.
/// </summary>
public class VariableWrapper
{
    private readonly WordprocessingDocumentWrapper _document;
    private readonly string _name;

    internal VariableWrapper(WordprocessingDocumentWrapper document, string name)
    {
        _document = document;
        _name = name;
    }

    private DocumentFormat.OpenXml.Wordprocessing.DocumentVariable Find()
    {
        return _document.GetSettingsPart(false)?.GetFirstChild<DocumentVariables>()?.Elements<DocumentFormat.OpenXml.Wordprocessing.DocumentVariable>()
            .FirstOrDefault(v => string.Equals(v.Name?.Value, _name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>The name of the variable.</summary>
    public string Name
    {
        get { return Find()?.Name?.Value ?? _name; }
    }

    /// <summary>The value of the variable (an empty string when the variable has been deleted).</summary>
    public string Value
    {
        get { return Find()?.Val?.Value ?? string.Empty; }
        set
        {
            DocumentFormat.OpenXml.Wordprocessing.DocumentVariable variable = Find();
            if (variable == null)
            {
                throw new InvalidOperationException("The variable has been deleted.");
            }
            variable.Val = value ?? string.Empty;
        }
    }

    /// <summary>Deletes the variable.</summary>
    public void Delete()
    {
        Find()?.Remove();
    }
}

/// <summary>
/// The variables of a document, as <c>Variables</c> in Word. A variable has a name and a text value, and can be shown with a DOCVARIABLE
/// field.
/// </summary>
public class VariableWrapperCollection : IEnumerable<VariableWrapper>
{
    private readonly WordprocessingDocumentWrapper _document;

    internal VariableWrapperCollection(WordprocessingDocumentWrapper document)
    {
        _document = document;
    }

    private IList<DocumentFormat.OpenXml.Wordprocessing.DocumentVariable> GetElementList()
    {
        return _document.GetSettingsPart(false)?.GetFirstChild<DocumentVariables>()?.Elements<DocumentFormat.OpenXml.Wordprocessing.DocumentVariable>().ToList() ?? new List<DocumentFormat.OpenXml.Wordprocessing.DocumentVariable>();
    }

    /// <summary>The number of variables.</summary>
    public int Count
    {
        get { return GetElementList().Count; }
    }

    /// <summary>The variable at the position.</summary>
    public VariableWrapper this[int index]
    {
        get
        {
            IList<DocumentFormat.OpenXml.Wordprocessing.DocumentVariable> variables = GetElementList();
            if (index < 0 || index >= variables.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return new VariableWrapper(_document, variables[index].Name.Value);
        }
    }

    /// <summary>The variable with the name (ignoring case), or null.</summary>
    public VariableWrapper this[string name]
    {
        get { return Exists(name) ? new VariableWrapper(_document, name) : null; }
    }

    /// <summary>True if there is a variable with the name (ignoring case).</summary>
    public bool Exists(string name)
    {
        return GetElementList().Any(v => string.Equals(v.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Adds a variable, or sets the value of the one with the name, and returns it.</summary>
    public VariableWrapper Add(string name, string value = "")
    {
        if (string.IsNullOrEmpty(name))
        {
            throw new ArgumentException("A variable needs a name.", nameof(name));
        }

        Settings settings = _document.GetSettingsPart(true);
        DocumentVariables variables = settings.GetFirstChild<DocumentVariables>();
        if (variables == null)
        {
            variables = new DocumentVariables();
            settings.AddChild(variables);
        }
        DocumentFormat.OpenXml.Wordprocessing.DocumentVariable existing = variables.Elements<DocumentFormat.OpenXml.Wordprocessing.DocumentVariable>().FirstOrDefault(v => string.Equals(v.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            existing.Val = value ?? string.Empty;
            return new VariableWrapper(_document, existing.Name.Value);
        }
        variables.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.DocumentVariable() { Name = name, Val = value ?? string.Empty });
        return new VariableWrapper(_document, name);
    }

    public IEnumerator<VariableWrapper> GetEnumerator()
    {
        foreach (DocumentFormat.OpenXml.Wordprocessing.DocumentVariable variable in GetElementList())
        {
            yield return new VariableWrapper(_document, variable.Name.Value);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
