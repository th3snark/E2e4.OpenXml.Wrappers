using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class ParagraphWrapper
{
    Paragraph _paragraph;

    public ParagraphWrapper(Paragraph paragraph)
    {
        _paragraph = paragraph;
    }

    public Paragraph Paragraph
    {
        get { return _paragraph; }
    }

    public RangeWrapper Range
    {
        get { return new RangeWrapper(_paragraph); }
    }

    public string Text
    {
        get { return Range.Text.Trim(new char[] { '\r' }); }
        set { Range.Text = value; }
    }

    public int SpaceBefore
    {
        get
        {
            if (_paragraph.ParagraphProperties != null)
            {
                if (_paragraph.ParagraphProperties.SpacingBetweenLines != null)
                {
                    return Convert.ToInt32(_paragraph.ParagraphProperties.SpacingBetweenLines.Before, CultureInfo.InvariantCulture);
                }
            }
            return 0;
        }

        set
        {
            if (_paragraph.ParagraphProperties == null)
            {
                _paragraph.ParagraphProperties = new ParagraphProperties();
            }
            if (_paragraph.ParagraphProperties.SpacingBetweenLines == null)
            {
                _paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines();
            }
            _paragraph.ParagraphProperties.SpacingBetweenLines.Before = value.ToString(CultureInfo.InvariantCulture);
        }
    }

    public int SpaceAfter
    {
        get
        {
            if (_paragraph.ParagraphProperties != null)
            {
                if (_paragraph.ParagraphProperties.SpacingBetweenLines != null)
                {
                    return Convert.ToInt32(_paragraph.ParagraphProperties.SpacingBetweenLines.After, CultureInfo.InvariantCulture);
                }
            }
            return 0;
        }

        set
        {
            if (_paragraph.ParagraphProperties == null)
            {
                _paragraph.ParagraphProperties = new ParagraphProperties();
            }
            if (_paragraph.ParagraphProperties.SpacingBetweenLines == null)
            {
                _paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines();
            }
            _paragraph.ParagraphProperties.SpacingBetweenLines.After = value.ToString(CultureInfo.InvariantCulture);
        }
    }

    public void Delete()
    {
        _paragraph.Parent.RemoveChild(_paragraph);
    }

    public void InsertBreak(BreakValues type)
    {
        if (type == BreakValues.Page || type == BreakValues.Column || type == BreakValues.TextWrapping)
        {
            _paragraph.AppendChild(new Run(new Break() { Type = type }));
        }
        else
        {
            throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "The break type '{0}' is not a recognised type.", type), nameof(type));
        }
    }

    public void DeleteBreaks()
    {
        foreach (Run run in _paragraph.Elements<Run>())
        {
            run.RemoveAllChildren<Break>();
        }
    }

    public SectionProperties InsertSectionBreak(string type)
    {
        switch (type)
        {
            case "continuous":
            case "evenPage":
            case "nextColumn":
            case "oddPage":
                throw new NotImplementedException(String.Format(CultureInfo.InvariantCulture, "The section break type '{0}' is not currrently supported.", type));

            case "nextPage":
                //Add a new paragraph properties if one does not exist.
                ParagraphProperties paragraphProperties = _paragraph.ParagraphProperties;
                if (paragraphProperties == null)
                {
                    paragraphProperties = new ParagraphProperties();
                    _paragraph.PrependChild(paragraphProperties);
                }

                //Remove the existing section properties from the paragraph properties.
                SectionProperties sectionProperties = paragraphProperties.SectionProperties;
                if (sectionProperties != null)
                {
                    sectionProperties.Parent.RemoveChild(sectionProperties);
                }

                //Get the last section properties in the body, clone it and then append it to the paragraph properties of this paragraph.
                SectionProperties finalSectionProperties = _paragraph.GetDocument().Body.Elements<SectionProperties>().First();
                SectionProperties newSectionProperties = (SectionProperties)finalSectionProperties.Clone();
                paragraphProperties.AppendChild(newSectionProperties);
                return newSectionProperties;

            default:
                throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "The section break type '{0}' is not a recognised type.", type), nameof(type));
        }
    }

    public void CopyPropertiesTo(ParagraphWrapper paragraphWrapper)
    {
        //Make the target paragraph properties the same as the source paragraph,
        ParagraphProperties fromParagraphProperties = (ParagraphProperties)_paragraph.ParagraphProperties;
        if (fromParagraphProperties != null)
        {
            if (paragraphWrapper.Paragraph.ParagraphProperties != null)
            {
                paragraphWrapper.Paragraph.RemoveChild(paragraphWrapper.Paragraph.ParagraphProperties);
            }
            ParagraphProperties toParagraphProperties = (ParagraphProperties)fromParagraphProperties.Clone();
            paragraphWrapper.Paragraph.PrependChild(toParagraphProperties);
        }

        //Make the all the target runs properties the same as the run properties of the first run of the source paragraph.
        Run firstRun = _paragraph.Elements<Run>().FirstOrDefault();
        if (firstRun != null)
        {
            RunProperties fromRunProperties = firstRun.RunProperties;
            if (fromRunProperties != null)
            {
                foreach (Run toRun in paragraphWrapper.Paragraph.Elements<Run>())
                {
                    RunProperties toRunProperties = toRun.RunProperties;
                    if (toRunProperties != null)
                    {
                        toRun.RemoveChild(toRunProperties);
                    }
                    toRun.PrependChild((RunProperties)fromRunProperties.Clone());
                }
            }
        }
    }

    public bool KeepNext
    {
        set
        {
            ParagraphProperties paragraphProperties = _paragraph.Elements<ParagraphProperties>().FirstOrDefault();
            if (paragraphProperties == null)
            {
                paragraphProperties = new ParagraphProperties();
                _paragraph.PrependChild<ParagraphProperties>(paragraphProperties);
            }
            ParagraphStyleId paragraphStyleId = paragraphProperties.Elements<ParagraphStyleId>().FirstOrDefault();
            KeepNext oldKeepNext = paragraphProperties.Elements<KeepNext>().FirstOrDefault();
            if (oldKeepNext != null)
            {
                if (oldKeepNext.HasAttributes)
                {
                    oldKeepNext.Val = value;
                }
                else
                {
                    if (!value)
                    {
                        oldKeepNext.Remove();
                    }
                }
            }
            else
            {
                if (value)
                {
                    KeepNext newKeepNext = new KeepNext();
                    if (paragraphStyleId != null)
                    {
                        paragraphStyleId.InsertAfterSelf<KeepNext>(newKeepNext);
                    }
                    else
                    {
                        paragraphProperties.PrependChild<KeepNext>(newKeepNext);
                    }
                }
            }
        }
    }

    public void SetFontBold()
    {
        SetFontBold(true);
    }

    public void SetFontBold(bool bold)
    {
        GetOrAddParagraphMarkRunProperties().SetFontBold(bold);

        foreach (Run run in _paragraph.Elements<Run>())
        {
            GetOrAddRunProperties(run).SetFontBold(bold);
        }
    }

    public void SetFontColor(string color)
    {
        GetOrAddParagraphMarkRunProperties().SetFontColor(color);

        foreach (Run run in _paragraph.Elements<Run>())
        {
            GetOrAddRunProperties(run).SetFontColor(color);
        }
    }

    private ParagraphMarkRunProperties GetOrAddParagraphMarkRunProperties()
    {
        ParagraphProperties paragraphProperties = _paragraph.ParagraphProperties;
        if (paragraphProperties == null)
        {
            paragraphProperties = new ParagraphProperties();
            _paragraph.PrependChild(paragraphProperties);
        }

        ParagraphMarkRunProperties paragraphMarkRunProperties = paragraphProperties.ParagraphMarkRunProperties;
        if (paragraphMarkRunProperties == null)
        {
            //The paragraph mark run properties must come before the section properties and the paragraph properties change.
            paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            OpenXmlElement successor = paragraphProperties.ChildElements.FirstOrDefault(e => e is SectionProperties || e is ParagraphPropertiesChange);
            if (successor != null)
            {
                successor.InsertBeforeSelf(paragraphMarkRunProperties);
            }
            else
            {
                paragraphProperties.AppendChild(paragraphMarkRunProperties);
            }
        }
        return paragraphMarkRunProperties;
    }

    private static RunProperties GetOrAddRunProperties(Run run)
    {
        RunProperties runProperties = run.RunProperties;
        if (runProperties == null)
        {
            runProperties = new RunProperties();
            run.PrependChild(runProperties);
        }
        return runProperties;
    }
}
