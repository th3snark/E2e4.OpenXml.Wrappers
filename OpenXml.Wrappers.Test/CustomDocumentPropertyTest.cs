using System;
using System.IO;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class CustomDocumentPropertyTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public CustomDocumentPropertyTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ADocumentStartsWithNoCustomPropertiesAndLookingDoesNotCreateThePart()
    {
        CustomDocumentPropertyCollection properties = _document.CustomDocumentProperties;

        Assert.Equal(0, properties.Count);
        Assert.Null(properties["Project"]);
        Assert.False(properties.Exists("Project"));
        Assert.Empty(properties);
        Assert.Null(_document.WordprocessingDocument.CustomFilePropertiesPart);
    }

    [Fact]
    public void AddStoresEachTypeAndReadsItBack()
    {
        CustomDocumentPropertyCollection properties = _document.CustomDocumentProperties;
        DateTime date = new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Utc);

        properties.Add("Project", DocumentPropertyType.Text, "Apollo");
        properties.Add("Revision", DocumentPropertyType.Number, 7);
        properties.Add("Ratio", DocumentPropertyType.FloatingPoint, 0.25);
        properties.Add("Approved", DocumentPropertyType.Boolean, true);
        properties.Add("Due", DocumentPropertyType.Date, date);

        Assert.Equal(5, properties.Count);
        Assert.Equal("Apollo", properties["Project"].Value);
        Assert.Equal(7, properties["Revision"].Value);
        Assert.Equal(0.25, properties["Ratio"].Value);
        Assert.Equal(true, properties["Approved"].Value);
        Assert.Equal(date, properties["Due"].Value);
        Assert.Equal(DocumentPropertyType.Date, properties["Due"].Type);
        Assert.Equal(DocumentPropertyType.Number, properties[1].Type);
    }

    [Fact]
    public void NamesAreCaseInsensitiveAndMustBeUnique()
    {
        CustomDocumentPropertyCollection properties = _document.CustomDocumentProperties;
        properties.Add("Project", DocumentPropertyType.Text, "Apollo");

        Assert.True(properties.Exists("PROJECT"));
        Assert.Equal("Project", properties["project"].Name);
        Assert.Throws<ArgumentException>(() => properties.Add("project", DocumentPropertyType.Text, "x"));
        Assert.Throws<ArgumentException>(() => properties.Add(" ", DocumentPropertyType.Text, "x"));
        Assert.Throws<ArgumentOutOfRangeException>(() => properties[1]);
    }

    [Fact]
    public void ValuesAreConvertedToTheTypeAndBadOnesRejected()
    {
        CustomDocumentPropertyCollection properties = _document.CustomDocumentProperties;

        properties.Add("N", DocumentPropertyType.Number, "12");
        properties.Add("S", DocumentPropertyType.Text, 5);
        properties.Add("Empty", DocumentPropertyType.Text, null);

        Assert.Equal(12, properties["N"].Value);
        Assert.Equal("5", properties["S"].Value);
        Assert.Equal(string.Empty, properties["Empty"].Value);
        Assert.Throws<ArgumentException>(() => properties.Add("Bad", DocumentPropertyType.Number, "abc"));
        Assert.False(properties.Exists("Bad"));
    }

    [Fact]
    public void IdsAreUniqueAndStartAtTwo()
    {
        CustomDocumentPropertyCollection properties = _document.CustomDocumentProperties;
        properties.Add("A", DocumentPropertyType.Text, "a");
        properties.Add("B", DocumentPropertyType.Text, "b");
        properties["A"].Delete();
        properties.Add("C", DocumentPropertyType.Text, "c");

        int[] ids = properties.Select(p => (int)p.Property.PropertyId.Value).ToArray();
        Assert.Equal(new[] { 3, 4 }, ids);
        Assert.All(properties, p => Assert.Equal("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.Property.FormatId.Value));
    }

    [Fact]
    public void SettingTheValueCanChangeTheType()
    {
        CustomDocumentPropertyCollection properties = _document.CustomDocumentProperties;
        CustomDocumentProperty property = properties.Add("P", DocumentPropertyType.Text, "one");

        property.Value = 2;
        Assert.Equal(DocumentPropertyType.Number, property.Type);
        Assert.Equal(2, property.Value);

        property.Value = "three";
        Assert.Equal(DocumentPropertyType.Text, property.Type);
        Assert.Single(property.Property.ChildElements);
        Assert.Throws<ArgumentException>(() => property.Value = new object());
    }

    [Fact]
    public void DeleteRemovesTheProperty()
    {
        CustomDocumentPropertyCollection properties = _document.CustomDocumentProperties;
        properties.Add("A", DocumentPropertyType.Text, "a").Delete();

        Assert.False(properties.Exists("A"));
        Assert.Equal(0, properties.Count);
    }

    [Fact]
    public void PropertiesSurviveSavingAndReopening()
    {
        _document.CustomDocumentProperties.Add("Project", DocumentPropertyType.Text, "Apollo");
        _document.CustomDocumentProperties.Add("Due", DocumentPropertyType.Date, new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc));
        string copy = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            _document.SaveAs(copy).Close();
            using (WordprocessingDocumentWrapper reopened = WordprocessingDocumentWrapper.Open(copy, false))
            {
                Assert.Equal("Apollo", reopened.CustomDocumentProperties["Project"].Value);
                Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc), reopened.CustomDocumentProperties["Due"].Value);
            }
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Fact]
    public void WrappersOfTheSamePropertyAreEqual()
    {
        _document.CustomDocumentProperties.Add("A", DocumentPropertyType.Text, "a");

        Assert.Equal(_document.CustomDocumentProperties["A"], _document.CustomDocumentProperties[0]);
    }
}
