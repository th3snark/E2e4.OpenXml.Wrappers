using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The document settings of Word's object model that are kept in <c>settings.xml</c>. Defaults and markup are Word 16's (the document-settings
/// scenario of MarkupConformanceTest).
/// </summary>
public class DocumentSettingsMembersTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public DocumentSettingsMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private string[] SettingNames()
    {
        return _document.WordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.Select(e => e.LocalName).ToArray();
    }

    [Fact]
    public void ANewDocumentHasWordsDefaults()
    {
        Assert.Equal(36, _document.DefaultTabStop);
        Assert.False(_document.AutoHyphenation);
        Assert.True(_document.HyphenateCaps);
        Assert.Equal(0, _document.ConsecutiveHyphensLimit);
        Assert.False(_document.UpdateStylesOnOpen);
        Assert.False(_document.EmbedTrueTypeFonts);
        Assert.False(_document.SaveSubsetFonts);
        Assert.True(_document.DoNotEmbedSystemFonts);
        Assert.False(_document.RemovePersonalInformation);
        Assert.False(_document.RemoveDateAndTime);
    }

    [Fact]
    public void TheSettingsReadBackAndAreWrittenInSchemaOrder()
    {
        _document.DefaultTabStop = 54;
        _document.AutoHyphenation = true;
        _document.HyphenateCaps = false;
        _document.ConsecutiveHyphensLimit = 2;
        _document.UpdateStylesOnOpen = true;
        _document.EmbedTrueTypeFonts = true;
        _document.SaveSubsetFonts = true;
        _document.DoNotEmbedSystemFonts = false;
        _document.RemovePersonalInformation = true;
        _document.RemoveDateAndTime = true;

        Assert.Equal(54, _document.DefaultTabStop);
        Assert.True(_document.AutoHyphenation);
        Assert.False(_document.HyphenateCaps);
        Assert.Equal(2, _document.ConsecutiveHyphensLimit);
        Assert.True(_document.UpdateStylesOnOpen);
        Assert.True(_document.EmbedTrueTypeFonts);
        Assert.True(_document.SaveSubsetFonts);
        Assert.False(_document.DoNotEmbedSystemFonts);
        Assert.True(_document.RemovePersonalInformation);
        Assert.True(_document.RemoveDateAndTime);

        string[] names = SettingNames();
        Assert.Equal(1, names.Count(n => n == "defaultTabStop"));
        Assert.True(Array.IndexOf(names, "removePersonalInformation") < Array.IndexOf(names, "embedTrueTypeFonts"));
        Assert.True(Array.IndexOf(names, "linkStyles") < Array.IndexOf(names, "defaultTabStop"));
        Assert.True(Array.IndexOf(names, "consecutiveHyphenLimit") < Array.IndexOf(names, "doNotHyphenateCaps"));
    }

    [Fact]
    public void ASettingThatIsBackToItsDefaultIsLeftOut()
    {
        _document.AutoHyphenation = true;
        _document.HyphenateCaps = false;
        _document.DoNotEmbedSystemFonts = false;
        _document.ConsecutiveHyphensLimit = 3;

        _document.AutoHyphenation = false;
        _document.HyphenateCaps = true;
        _document.DoNotEmbedSystemFonts = true;
        _document.ConsecutiveHyphensLimit = 0;

        string[] names = SettingNames();
        Assert.DoesNotContain("autoHyphenation", names);
        Assert.DoesNotContain("doNotHyphenateCaps", names);
        Assert.DoesNotContain("embedSystemFonts", names);
        Assert.DoesNotContain("consecutiveHyphenLimit", names);
    }

    [Fact]
    public void ValuesOutOfRangeAreRejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.DefaultTabStop = 0);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.DefaultTabStop = 2000);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.ConsecutiveHyphensLimit = -1);
        Assert.Equal(36, _document.DefaultTabStop);
    }

    [Fact]
    public void ADocumentIsADocumentAndItsFormatIsDocx()
    {
        Assert.Equal(WordDocumentType.Document, _document.Type);
        Assert.Equal(SaveFormat.XmlDocument, _document.SaveFormat);

        _document.WordprocessingDocument.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Template);

        Assert.Equal(WordDocumentType.Template, _document.Type);
        Assert.Equal(SaveFormat.XmlTemplate, _document.SaveFormat);
    }

    [Fact]
    public void DeleteAllCommentsRemovesEveryComment()
    {
        _document.Content.InsertAfter("Some text");
        _document.Comments.Add(_document.Range(0, 4), "one");
        _document.Comments.Add(_document.Range(5, 9), "two");
        Assert.Equal(2, _document.Comments.Count);

        _document.DeleteAllComments();

        Assert.Equal(0, _document.Comments.Count);
        Assert.Equal("Some text", _document.Content.Text.TrimEnd('\r'));
    }

    [Fact]
    public void TheCompatibilityModeCanBeSetByAMethod()
    {
        _document.SetCompatibilityMode(14);

        Assert.Equal(14, _document.CompatibilityMode);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.SetCompatibilityMode(13));
    }
}
