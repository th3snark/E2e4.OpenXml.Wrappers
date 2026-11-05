using System;
using System.Collections.Generic;
using System.Text;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

internal sealed partial class StoryMap
{
    private const char FieldBeginChar = '\u0013';
    private const char FieldEndChar = '\u0015';

    /// <summary>
    /// The text of <c>Range.Text</c> as <see cref="TextRetrievalMode"/> sets it, for the positions from <paramref name="start"/> up to
    /// <paramref name="end"/>. With field codes, a field shows U+0013, its code and U+0015 and leaves out its separator and result; without,
    /// it shows its result only, as it always does. Hidden text is left out unless <paramref name="includeHiddenText"/>.
    /// </summary>
    public string GetText(int start, int end, bool includeFieldCodes, bool includeHiddenText)
    {
        StringBuilder sb = new StringBuilder();
        FormattingResolver resolver = includeHiddenText ? null : new FormattingResolver(Root);

        //For each field that is open, innermost last: true while it is in its code, false while it is in its result.
        List<bool> inCode = new List<bool>();
        foreach (Atom atom in Atoms)
        {
            if (atom.Start >= end)
            {
                break;
            }

            //Codes show text that is in the codes of fields only; results show text that is not in any code.
            bool visible = Shows(inCode, includeFieldCodes);
            bool inRange = atom.End > start;
            int from = Math.Max(start, atom.Start) - atom.Start;
            int to = Math.Min(end, atom.End) - atom.Start;

            if (atom.Kind == AtomKind.FieldMark && atom.Element is FieldChar mark)
            {
                FieldCharValues type = mark.FieldCharType?.Value ?? FieldCharValues.Begin;
                if (type == FieldCharValues.Begin)
                {
                    AppendIf(sb, inRange && includeFieldCodes && visible, FieldBeginChar);
                    inCode.Add(true);
                }
                else if (type == FieldCharValues.Separate)
                {
                    if (inCode.Count > 0)
                    {
                        inCode[inCode.Count - 1] = false;
                    }
                }
                else
                {
                    if (inCode.Count > 0)
                    {
                        inCode.RemoveAt(inCode.Count - 1);
                    }
                    AppendIf(sb, inRange && includeFieldCodes && Shows(inCode, true), FieldEndChar);
                }
            }
            else if (atom.Kind == AtomKind.FieldCode && atom.IsSimpleFieldPart)
            {
                //A field written as one element: this atom is its begin, code and separator, and its end too when it has no result.
                string code = atom.Element is SimpleField simple ? simple.Instruction?.Value ?? string.Empty : HyperlinkCode((Hyperlink)atom.Element);
                bool hasResult = atom.Length == code.Length + 2;
                if (inRange && includeFieldCodes && visible)
                {
                    string shown = FieldBeginChar + code + "\0" + (hasResult ? string.Empty : FieldEndChar.ToString());
                    for (int i = from; i < to; i++)
                    {
                        if (shown[i] != '\0')
                        {
                            sb.Append(shown[i]);
                        }
                    }
                }
                if (hasResult)
                {
                    inCode.Add(false);
                }
            }
            else if (atom.Kind == AtomKind.FieldMark)
            {
                //The end of a field written as one element.
                if (inCode.Count > 0)
                {
                    inCode.RemoveAt(inCode.Count - 1);
                }
                AppendIf(sb, inRange && includeFieldCodes && Shows(inCode, true), FieldEndChar);
            }
            else if (atom.Kind == AtomKind.FieldCode)
            {
                if (inRange && includeFieldCodes && visible && atom.Element is FieldCode complexCode && complexCode.Text != null)
                {
                    sb.Append(complexCode.Text, from, to - from);
                }
            }
            else if (inRange && visible && !IsHiddenText(resolver, atom))
            {
                if (atom.Kind == AtomKind.Text)
                {
                    sb.Append(atom.Text, from, to - from);
                }
                else
                {
                    sb.Append(atom.Text);
                }
            }
        }
        return sb.ToString();
    }

    //True if text is shown while the fields in the list are open: with codes, only when all of them are in their code; without, when none is.
    private static bool Shows(List<bool> inCode, bool includeFieldCodes)
    {
        return includeFieldCodes ? !inCode.Contains(false) : !inCode.Contains(true);
    }

    private static void AppendIf(StringBuilder sb, bool condition, char value)
    {
        if (condition)
        {
            sb.Append(value);
        }
    }

    private static bool IsHiddenText(FormattingResolver resolver, Atom atom)
    {
        if (resolver == null || atom.IsMark || !(atom.Element?.Parent is Run run))
        {
            return false;
        }
        Vanish vanish = resolver.ResolveRun(run)?.GetFirstChild<Vanish>();
        return vanish != null && (vanish.Val == null || vanish.Val.Value);
    }
}
