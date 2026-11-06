using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The members of content controls that belong to one kind of control. The markup is Word 16's (the content-control-* scenarios of
/// MarkupConformanceTest).
/// </summary>
public class ContentControlMembersTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ContentControlMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Hello world again");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        System.Collections.Generic.IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    [Fact]
    public void ACheckBoxIsTickedAndShowsTheSymbolOfItsState()
    {
        ContentControlWrapper box = _document.ContentControls.Add(ContentControlType.CheckBox, _document.Range(6, 6));

        Assert.False(box.Checked);
        box.Checked = true;
        Assert.True(box.Checked);
        Assert.Equal("☒", box.Range.Text);
        box.Checked = false;
        Assert.Equal("☐", box.Range.Text);

        box.SetCheckedSymbol(10004, "Segoe UI Symbol");
        box.SetUncheckedSymbol(10006);
        Assert.Equal("✖", box.Range.Text);
        box.Checked = true;
        Assert.Equal("✔", box.Range.Text);
        AssertValid();
    }

    [Fact]
    public void ListEntriesAreAddedMovedChosenAndCleared()
    {
        ContentControlWrapper list = _document.ContentControls.Add(ContentControlType.DropdownList, _document.Range(5, 5));
        ContentControlListEntries entries = list.DropdownListEntries;

        ContentControlListEntry one = entries.Add("One", "1");
        entries.Add("Two");
        entries.Add("Zero", "0", 0);

        Assert.Equal(3, entries.Count);
        Assert.Equal(new[] { "Zero", "One", "Two" }, entries.Select(e => e.Text).ToArray());
        Assert.Equal("Two", entries[2].Value);
        Assert.Equal(1, one.Index);

        one.MoveDown();
        Assert.Equal(new[] { "Zero", "Two", "One" }, entries.Select(e => e.Text).ToArray());
        one.MoveUp();
        one.MoveUp();
        Assert.Equal(0, one.Index);
        one.MoveUp();
        Assert.Equal(0, one.Index);

        Assert.True(list.ShowingPlaceholderText);
        entries[1].Select();
        Assert.False(list.ShowingPlaceholderText);
        Assert.Equal("Zero", list.Range.Text);

        entries[0].Delete();
        Assert.Equal(2, entries.Count);
        entries.Clear();
        Assert.Equal(0, entries.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => entries[0]);
        Assert.Throws<ArgumentOutOfRangeException>(() => entries.Add("x", null, 1));
        AssertValid();
    }

    [Fact]
    public void ListEntriesBelongToListsAndComboBoxesOnly()
    {
        ContentControlWrapper combo = _document.ContentControls.Add(ContentControlType.ComboBox, _document.Range(0, 0));
        combo.DropdownListEntries.Add("A", "a");
        ContentControlWrapper text = _document.ContentControls.Add(ContentControlType.Text, _document.Range(0, 0));

        Assert.Equal(1, combo.DropdownListEntries.Count);
        Assert.Throws<InvalidOperationException>(() => text.DropdownListEntries);
        Assert.Throws<InvalidOperationException>(() => text.Checked);
        Assert.Throws<InvalidOperationException>(() => combo.MultiLine);
        Assert.Throws<InvalidOperationException>(() => combo.DateDisplayFormat);
    }

    [Fact]
    public void ADatePickerHasAFormatALocaleACalendarAndAStorageFormat()
    {
        ContentControlWrapper date = _document.ContentControls.Add(ContentControlType.Date, _document.Range(5, 5));

        Assert.Equal(ContentControlDateStorageFormat.DateTime, date.DateStorageFormat);
        Assert.Equal(CalendarType.Western, date.DateCalendarType);

        date.DateDisplayFormat = "dddd, d MMMM yyyy";
        date.DateStorageFormat = ContentControlDateStorageFormat.Date;
        date.DateDisplayLocale = 1041;
        date.DateCalendarType = CalendarType.Japan;

        Assert.Equal("dddd, d MMMM yyyy", date.DateDisplayFormat);
        Assert.Equal(ContentControlDateStorageFormat.Date, date.DateStorageFormat);
        Assert.Equal(1041, date.DateDisplayLocale);
        Assert.Equal(CalendarType.Japan, date.DateCalendarType);
        Assert.Throws<ArgumentOutOfRangeException>(() => date.DateCalendarType = (CalendarType)99);
        AssertValid();
    }

    [Fact]
    public void ATextControlCanRunOverSeveralLines()
    {
        ContentControlWrapper text = _document.ContentControls.Add(ContentControlType.Text, _document.Range(5, 5));

        Assert.False(text.MultiLine);
        text.MultiLine = true;
        Assert.True(text.MultiLine);
        text.MultiLine = false;
        Assert.False(text.MultiLine);
        AssertValid();
    }

    [Fact]
    public void AControlHasAColourALevelAndATemporaryFlag()
    {
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.RichText, _document.Range(6, 11));

        Assert.True(control.Color.IsAutomatic);
        Assert.False(control.Temporary);
        Assert.Equal(ContentControlLevel.Inline, control.Level);

        control.Color = 255;
        control.Temporary = true;

        Assert.Equal(255, control.Color.Value);
        Assert.True(control.Temporary);

        control.Color = WordColor.Automatic;
        control.Temporary = false;
        Assert.True(control.Color.IsAutomatic);
        Assert.False(control.Temporary);
        AssertValid();
    }

    [Fact]
    public void TheDefaultTextStyleIsAppliedToTheContents()
    {
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(6, 11));

        Assert.Null(control.DefaultTextStyle);
        control.DefaultTextStyle = "Strong";

        Assert.Equal("Strong", control.DefaultTextStyle);
        Assert.Equal("Strong", control.Range.Style);
        Assert.Throws<ArgumentException>(() => control.DefaultTextStyle = "Heading 1");

        control.DefaultTextStyle = null;
        Assert.Null(control.DefaultTextStyle);
        AssertValid();
    }
}
