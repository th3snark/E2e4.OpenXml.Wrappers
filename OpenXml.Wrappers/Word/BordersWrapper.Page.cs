using System;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Where the distance of a page border is measured from, as <c>WdPageBorderDistanceFrom</c> in Word (with the same values).
/// </summary>
public enum BorderDistanceFrom
{
    /// <summary>From the text.</summary>
    FromText = 0,

    /// <summary>From the edge of the page.</summary>
    FromPageEdge = 1
}

public partial class BordersWrapper
{
    private PageBorders GetPageBorders(bool create)
    {
        if (!(_owner is SectionProperties))
        {
            throw new NotSupportedException("Only the page borders of a section have this member.");
        }
        return (PageBorders)GetContainer(create);
    }

    /// <summary>
    /// True (the default) to draw the page border in front of the text, false for behind it, as <c>Borders.AlwaysInFront</c> in Word. Only the
    /// page borders of a section have it.
    /// </summary>
    /// <exception cref="NotSupportedException">The borders are not those of a page.</exception>
    public bool AlwaysInFront
    {
        get { return GetPageBorders(false)?.ZOrder?.Value != PageBorderZOrderValues.Back; }
        set
        {
            PageBorders borders = GetPageBorders(true);
            borders.ZOrder = value ? null : PageBorderZOrderValues.Back;
        }
    }

    /// <summary>
    /// Where the distance of the page border is measured from, as <c>Borders.DistanceFrom</c> in Word. The page borders of a new section are
    /// measured from the edge of the page.
    /// </summary>
    /// <exception cref="NotSupportedException">The borders are not those of a page.</exception>
    public BorderDistanceFrom DistanceFrom
    {
        get
        {
            PageBorders borders = GetPageBorders(false);
            return borders == null || borders.OffsetFrom?.Value == PageBorderOffsetValues.Page ? BorderDistanceFrom.FromPageEdge : BorderDistanceFrom.FromText;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            //The schema's default is the text, which Word does not write.
            GetPageBorders(true).OffsetFrom = value == BorderDistanceFrom.FromPageEdge ? PageBorderOffsetValues.Page : null;
        }
    }

    private PageBorderDisplayValues? GetDisplay()
    {
        return GetPageBorders(false)?.Display?.Value;
    }

    /// <summary>
    /// True (the default) to show the page border on the first page of the section, as <c>Borders.EnableFirstPageInSection</c> in Word. As in
    /// Word, at least one of this and <see cref="EnableOtherPagesInSection"/> stays true: turning one off when the other is off turns the other on.
    /// </summary>
    /// <exception cref="NotSupportedException">The borders are not those of a page.</exception>
    public bool EnableFirstPageInSection
    {
        get { return GetDisplay() != PageBorderDisplayValues.NotFirstPage; }
        set { SetDisplay(value, EnableOtherPagesInSection); }
    }

    /// <summary>
    /// True (the default) to show the page border on the pages of the section after the first, as <c>Borders.EnableOtherPagesInSection</c> in Word.
    /// See <see cref="EnableFirstPageInSection"/>.
    /// </summary>
    /// <exception cref="NotSupportedException">The borders are not those of a page.</exception>
    public bool EnableOtherPagesInSection
    {
        get { return GetDisplay() != PageBorderDisplayValues.FirstPage; }
        set { SetDisplay(EnableFirstPageInSection, value); }
    }

    private void SetDisplay(bool first, bool other)
    {
        PageBorders borders = GetPageBorders(true);
        //Switching one off while the other is off turns the other on: the page borders cannot be on no page.
        if (!first && !other)
        {
            first = !EnableFirstPageInSection;
            other = !EnableOtherPagesInSection;
        }
        borders.Display = first && other ? null : first ? PageBorderDisplayValues.FirstPage : PageBorderDisplayValues.NotFirstPage;
    }

    //Word writes a setting for the header and footer to be outside the page border, and nothing when they are inside it.
    private bool IsOutside<T>() where T : OnOffType
    {
        GetPageBorders(false);
        T setting = GetSettings(false)?.GetFirstChild<T>();
        return setting != null && (setting.Val == null || !setting.Val.HasValue || setting.Val.Value);
    }

    private void SetSurround<T>(bool value) where T : OnOffType, new()
    {
        GetPageBorders(false);
        if (value)
        {
            GetSettings(false)?.RemoveAllChildren<T>();
        }
        else if (!IsOutside<T>())
        {
            Settings settings = GetSettings(true);
            settings.RemoveAllChildren<T>();
            settings.AddChild(new T());
        }
    }

    private Settings GetSettings(bool create)
    {
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(_owner);
        if (main == null)
        {
            if (create)
            {
                throw new InvalidOperationException("The section is not part of a document.");
            }
            return null;
        }
        if (main.DocumentSettingsPart == null && create)
        {
            main.AddNewPart<DocumentSettingsPart>();
        }
        if (main.DocumentSettingsPart != null && main.DocumentSettingsPart.Settings == null && create)
        {
            main.DocumentSettingsPart.Settings = new Settings();
        }
        return main.DocumentSettingsPart?.Settings;
    }

    /// <summary>
    /// True (the default) to put the header inside the page border, as <c>Borders.SurroundHeader</c> in Word. It is a setting of the document,
    /// kept in its settings.
    /// </summary>
    /// <exception cref="NotSupportedException">The borders are not those of a page.</exception>
    public bool SurroundHeader
    {
        get { return !IsOutside<BordersDoNotSurroundHeader>(); }
        set { SetSurround<BordersDoNotSurroundHeader>(value); }
    }

    /// <summary>True (the default) to put the footer inside the page border, as <c>Borders.SurroundFooter</c> in Word. See <see cref="SurroundHeader"/>.</summary>
    /// <exception cref="NotSupportedException">The borders are not those of a page.</exception>
    public bool SurroundFooter
    {
        get { return !IsOutside<BordersDoNotSurroundFooter>(); }
        set { SetSurround<BordersDoNotSurroundFooter>(value); }
    }
}
