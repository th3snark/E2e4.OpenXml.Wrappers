using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The page setup of a section, as <c>PageSetup</c> in Word. Measurements are in points.
/// </summary>
/// <remarks>
/// A section without a setting reads the default of Word: a Letter page (612 x 792 points) with 72 point margins and 36 point header and
/// footer distances, that starts on a new page. Setting a value writes it into the section properties of the section.
/// </remarks>
public class PageSetupWrapper
{
    private const double MaximumPoints = 1584;

    //The order of the child elements of the section properties in the schema.
    private static readonly string[] SchemaOrder =
    {
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
        "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings"
    };

    private readonly SectionWrapper _section;

    internal PageSetupWrapper(SectionWrapper section)
    {
        _section = section;
    }

    private SectionProperties Properties
    {
        get { return _section.GetSectionProperties(false); }
    }

    private T GetOrAdd<T>() where T : OpenXmlElement, new()
    {
        SectionProperties properties = _section.GetSectionProperties(true);
        T element = properties.GetFirstChild<T>();
        if (element == null)
        {
            element = new T();
            int rank = Array.IndexOf(SchemaOrder, element.LocalName);
            OpenXmlElement next = properties.ChildElements.FirstOrDefault(e => Array.IndexOf(SchemaOrder, e.LocalName) > rank);
            if (next != null)
            {
                next.InsertBeforeSelf(element);
            }
            else
            {
                properties.AppendChild(element);
            }
        }
        return element;
    }

    private static void CheckPoints(double value, string parameterName)
    {
        if (Double.IsNaN(value) || value < 0 || value > MaximumPoints)
        {
            throw new ArgumentOutOfRangeException(parameterName, String.Format(CultureInfo.InvariantCulture, "The value must be between 0 and {0} points.", MaximumPoints));
        }
    }

    private static uint ToTwips(double points)
    {
        return (uint)Math.Round(points * 20, MidpointRounding.AwayFromZero);
    }

    /// <summary>
    /// The width of the page in points.
    /// </summary>
    public double PageWidth
    {
        get { return (Properties?.GetFirstChild<PageSize>()?.Width?.Value ?? 12240u) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageSize>().Width = ToTwips(value);
        }
    }

    /// <summary>
    /// The height of the page in points.
    /// </summary>
    public double PageHeight
    {
        get { return (Properties?.GetFirstChild<PageSize>()?.Height?.Value ?? 15840u) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageSize>().Height = ToTwips(value);
        }
    }

    /// <summary>
    /// The orientation of the page. Setting it swaps the width and height of the page when the orientation changes, as Word does.
    /// </summary>
    public PageOrientation Orientation
    {
        get
        {
            PageSize pageSize = Properties?.GetFirstChild<PageSize>();
            if (pageSize?.Orient != null && pageSize.Orient.HasValue)
            {
                return pageSize.Orient.Value == PageOrientationValues.Landscape ? PageOrientation.Landscape : PageOrientation.Portrait;
            }
            return PageWidth > PageHeight ? PageOrientation.Landscape : PageOrientation.Portrait;
        }
        set
        {
            if (!Enum.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            bool changed = value != Orientation;
            double width = PageWidth;
            double height = PageHeight;
            PageSize pageSize = GetOrAdd<PageSize>();
            //A page that has no size yet is given the default one so that it can be swapped.
            pageSize.Width = ToTwips(changed ? height : width);
            pageSize.Height = ToTwips(changed ? width : height);
            pageSize.Orient = value == PageOrientation.Landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
        }
    }

    /// <summary>
    /// The distance between the top of the page and the text, in points.
    /// </summary>
    public double TopMargin
    {
        get { return (Properties?.GetFirstChild<PageMargin>()?.Top?.Value ?? 1440) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageMargin>().Top = (int)ToTwips(value);
        }
    }

    /// <summary>
    /// The distance between the bottom of the page and the text, in points.
    /// </summary>
    public double BottomMargin
    {
        get { return (Properties?.GetFirstChild<PageMargin>()?.Bottom?.Value ?? 1440) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageMargin>().Bottom = (int)ToTwips(value);
        }
    }

    /// <summary>
    /// The distance between the left edge of the page and the text, in points.
    /// </summary>
    public double LeftMargin
    {
        get { return (Properties?.GetFirstChild<PageMargin>()?.Left?.Value ?? 1440u) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageMargin>().Left = ToTwips(value);
        }
    }

    /// <summary>
    /// The distance between the right edge of the page and the text, in points.
    /// </summary>
    public double RightMargin
    {
        get { return (Properties?.GetFirstChild<PageMargin>()?.Right?.Value ?? 1440u) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageMargin>().Right = ToTwips(value);
        }
    }

    /// <summary>
    /// The extra space for binding, in points.
    /// </summary>
    public double Gutter
    {
        get { return (Properties?.GetFirstChild<PageMargin>()?.Gutter?.Value ?? 0u) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageMargin>().Gutter = ToTwips(value);
        }
    }

    /// <summary>
    /// The distance between the top of the page and the header, in points.
    /// </summary>
    public double HeaderDistance
    {
        get { return (Properties?.GetFirstChild<PageMargin>()?.Header?.Value ?? 720u) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageMargin>().Header = ToTwips(value);
        }
    }

    /// <summary>
    /// The distance between the bottom of the page and the footer, in points.
    /// </summary>
    public double FooterDistance
    {
        get { return (Properties?.GetFirstChild<PageMargin>()?.Footer?.Value ?? 720u) / 20.0; }
        set
        {
            CheckPoints(value, nameof(value));
            GetOrAdd<PageMargin>().Footer = ToTwips(value);
        }
    }

    /// <summary>
    /// Where the section starts. The default is <see cref="Word.SectionStart.NewPage"/>.
    /// </summary>
    public SectionStart SectionStart
    {
        get
        {
            SectionMarkValues? type = Properties?.GetFirstChild<SectionType>()?.Val?.Value;
            if (type == SectionMarkValues.Continuous)
            {
                return SectionStart.Continuous;
            }
            if (type == SectionMarkValues.NextColumn)
            {
                return SectionStart.NewColumn;
            }
            if (type == SectionMarkValues.EvenPage)
            {
                return SectionStart.EvenPage;
            }
            return type == SectionMarkValues.OddPage ? SectionStart.OddPage : SectionStart.NewPage;
        }
        set
        {
            if (!Enum.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            SectionMarkValues type = value switch
            {
                SectionStart.Continuous => SectionMarkValues.Continuous,
                SectionStart.NewColumn => SectionMarkValues.NextColumn,
                SectionStart.EvenPage => SectionMarkValues.EvenPage,
                SectionStart.OddPage => SectionMarkValues.OddPage,
                _ => SectionMarkValues.NextPage
            };
            GetOrAdd<SectionType>().Val = type;
        }
    }

    /// <summary>
    /// True when the first page of the section has a different header and footer, as <c>PageSetup.DifferentFirstPageHeaderFooter</c> in Word.
    /// </summary>
    public bool DifferentFirstPageHeaderFooter
    {
        get
        {
            TitlePage titlePage = Properties?.GetFirstChild<TitlePage>();
            return titlePage != null && (titlePage.Val == null || !titlePage.Val.HasValue || titlePage.Val.Value);
        }
        set
        {
            if (value)
            {
                GetOrAdd<TitlePage>().Val = null;
            }
            else
            {
                Properties?.RemoveAllChildren<TitlePage>();
            }
        }
    }
}
