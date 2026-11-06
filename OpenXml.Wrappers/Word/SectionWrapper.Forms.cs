using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class SectionWrapper
{
    /// <summary>
    /// True (the default) if the section is protected when the document is protected for forms, as <c>Section.ProtectedForForms</c> in Word;
    /// false for a section that stays editable. It has an effect only while <see cref="WordprocessingDocumentWrapper.Protect"/> has protected the
    /// document for form fields.
    /// </summary>
    public bool ProtectedForForms
    {
        get
        {
            FormProtection protection = SectionProperties?.GetFirstChild<FormProtection>();
            return protection == null || protection.Val == null || !protection.Val.HasValue || protection.Val.Value;
        }
        set
        {
            if (value)
            {
                SectionProperties?.RemoveAllChildren<FormProtection>();
                return;
            }

            SectionProperties properties = GetSectionProperties(true);
            properties.RemoveAllChildren<FormProtection>();
            //Word writes 0 where the SDK would write false.
            FormProtection off = new FormProtection();
            off.SetAttribute(new OpenXmlAttribute("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0"));
            properties.AddChild(off);
        }
    }
}
