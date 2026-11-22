using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Works out a formula (<c>=</c>) field that does not need a table: numbers, arithmetic (<c>+ - * / ^ %</c>), comparisons (which give 1 or
/// 0), Word's functions and bookmarks that hold a number, as measured in Word. Formulas whose result would depend on things not worked
/// out here (cell references, numbers written with digit grouping, very large numbers and syntax errors, whose messages Word words in its own
/// way) give null, which leaves the field as it is.
/// </summary>
internal sealed class FieldFormulas
{
    private sealed class CannotComputeException : Exception
    {
    }

    private sealed class FormulaErrorException : Exception
    {
        public FormulaErrorException(string message)
            : base(message)
        {
        }
    }

    private static readonly string[] Directions = { "ABOVE", "BELOW", "LEFT", "RIGHT" };

    private readonly string _text;
    private readonly OpenXmlElement _root;
    private readonly bool _inTable;
    private readonly NumberFormatInfo _numbers;
    private readonly string _listSeparator;
    private int _position;

    //The most decimals of a number in the formula, which the result is shown with (and at least 2).
    private int _decimals;

    private FieldFormulas(string text, OpenXmlElement root, bool inTable)
    {
        _text = text;
        _root = root;
        _inTable = inTable;
        _numbers = CultureInfo.CurrentCulture.NumberFormat;
        _listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator.Trim();
    }

    /// <summary>
    /// The result of the formula (the text after the <c>=</c>, without its switches), or null when it is not one that can be worked out here.
    /// </summary>
    internal static string Compute(string expression, OpenXmlElement root, bool inTable, string picture)
    {
        FieldFormulas formulas = new FieldFormulas(expression, root, inTable);
        if (formulas._listSeparator.Length != 1 || formulas._listSeparator == formulas._numbers.NumberDecimalSeparator)
        {
            return null;
        }

        double value;
        try
        {
            value = formulas.Evaluate();
        }
        catch (FormulaErrorException error)
        {
            return error.Message;
        }
        catch (CannotComputeException)
        {
            return null;
        }

        if (double.IsNaN(value) || double.IsInfinity(value) || Math.Abs(value) >= 1e14)
        {
            return null;
        }
        if (picture != null)
        {
            try
            {
                return value.ToString(picture, CultureInfo.CurrentCulture);
            }
            catch (FormatException)
            {
                return null;
            }
        }
        int decimals = Math.Max(2, formulas._decimals);
        return Math.Round(value, Math.Min(decimals, 15), MidpointRounding.AwayFromZero).ToString("0." + new string('#', decimals), CultureInfo.CurrentCulture);
    }

    private double Evaluate()
    {
        SkipSpaces();
        if (AtEnd())
        {
            throw new FormulaErrorException(TableFormulas.UnexpectedEnd);
        }
        double value = Comparison();
        SkipSpaces();
        if (!AtEnd())
        {
            throw new CannotComputeException();
        }
        return value;
    }

    #region Grammar

    private double Comparison()
    {
        double left = Additive();
        SkipSpaces();
        foreach (string op in new[] { "<=", ">=", "<>", "=", "<", ">" })
        {
            if (string.CompareOrdinal(_text, _position, op, 0, op.Length) == 0)
            {
                _position += op.Length;
                double right = Additive();
                bool result = op == "<=" ? left <= right : op == ">=" ? left >= right : op == "<>" ? left != right
                    : op == "=" ? left == right : op == "<" ? left < right : left > right;
                return result ? 1 : 0;
            }
        }
        return left;
    }

    private double Additive()
    {
        double value = Term();
        while (true)
        {
            SkipSpaces();
            if (Accept('+'))
            {
                value += Term();
            }
            else if (Accept('-'))
            {
                value -= Term();
            }
            else
            {
                return value;
            }
        }
    }

    private double Term()
    {
        double value = Power();
        while (true)
        {
            SkipSpaces();
            if (Accept('*'))
            {
                value *= Power();
            }
            else if (Accept('/'))
            {
                double divisor = Power();
                if (divisor == 0)
                {
                    throw new FormulaErrorException("!Zero Divide");
                }
                value /= divisor;
            }
            else
            {
                return value;
            }
        }
    }

    private double Power()
    {
        double value = Unary();
        SkipSpaces();
        while (Accept('^'))
        {
            value = Math.Pow(value, Unary());
            SkipSpaces();
        }
        return value;
    }

    private double Unary()
    {
        SkipSpaces();
        if (Accept('-'))
        {
            return -Unary();
        }
        if (Accept('+'))
        {
            return Unary();
        }
        double value = Primary();
        SkipSpaces();
        if (Accept('%'))
        {
            value /= 100;
        }
        return value;
    }

    private double Primary()
    {
        SkipSpaces();
        if (AtEnd())
        {
            throw new FormulaErrorException(TableFormulas.UnexpectedEnd);
        }
        char c = _text[_position];
        if (c == '(')
        {
            _position++;
            double value = Comparison();
            SkipSpaces();
            Expect(')');
            return value;
        }
        if (char.IsDigit(c) || _text.Substring(_position).StartsWith(_numbers.NumberDecimalSeparator, StringComparison.Ordinal))
        {
            return Number();
        }
        if (char.IsLetter(c) || c == '_')
        {
            int start = _position;
            while (!AtEnd() && (char.IsLetterOrDigit(_text[_position]) || _text[_position] == '_'))
            {
                _position++;
            }
            string name = _text.Substring(start, _position - start);
            SkipSpaces();
            if (Accept('('))
            {
                return Function(name.ToUpperInvariant());
            }
            return Name(name);
        }
        throw new CannotComputeException();
    }

    private double Number()
    {
        int start = _position;
        while (!AtEnd() && char.IsDigit(_text[_position]))
        {
            _position++;
        }
        int decimals = 0;
        if (string.CompareOrdinal(_text, _position, _numbers.NumberDecimalSeparator, 0, _numbers.NumberDecimalSeparator.Length) == 0)
        {
            _position += _numbers.NumberDecimalSeparator.Length;
            int fraction = _position;
            while (!AtEnd() && char.IsDigit(_text[_position]))
            {
                _position++;
            }
            decimals = _position - fraction;
        }
        //1e3, or a digit grouping such as 1,234, is not a number to Word.
        if (!AtEnd() && (char.IsLetter(_text[_position]) || _text[_position] == '_'))
        {
            throw new CannotComputeException();
        }
        _decimals = Math.Max(_decimals, decimals);
        string digits = _text.Substring(start, _position - start);
        return double.Parse(digits, NumberStyles.AllowDecimalPoint, _numbers);
    }

    private List<double> Arguments()
    {
        List<double> values = new List<double>();
        SkipSpaces();
        if (Accept(')'))
        {
            return values;
        }
        while (true)
        {
            values.Add(Comparison());
            SkipSpaces();
            if (Accept(')'))
            {
                return values;
            }
            Expect(_listSeparator[0]);
        }
    }

    private double Function(string name)
    {
        if (name == "DEFINED")
        {
            SkipSpaces();
            int start = _position;
            while (!AtEnd() && _text[_position] != ')')
            {
                _position++;
            }
            string bookmark = _text.Substring(start, _position - start).Trim();
            Expect(')');
            return FindBookmark(bookmark) != null ? 1 : 0;
        }

        SkipSpaces();
        foreach (string direction in Directions)
        {
            if (string.Compare(_text, _position, direction, 0, direction.Length, StringComparison.OrdinalIgnoreCase) == 0)
            {
                //A function of the cells around is worked out by TableFormulas in a table, and has no table to work on elsewhere.
                if (_inTable)
                {
                    throw new CannotComputeException();
                }
                throw new FormulaErrorException("!The Formula Not In Table");
            }
        }

        List<double> values = Arguments();
        switch (name)
        {
            case "ABS":
                return Math.Abs(Single(values));
            case "INT":
                return Math.Truncate(Single(values));
            case "SIGN":
                return Math.Sign(Single(values));
            case "NOT":
                return Single(values) == 0 ? 1 : 0;
            case "TRUE":
                Count(values, 0);
                return 1;
            case "FALSE":
                Count(values, 0);
                return 0;
            case "MOD":
                Count(values, 2);
                if (values[1] == 0)
                {
                    throw new FormulaErrorException("!Zero Divide");
                }
                return values[0] % values[1];
            case "ROUND":
                Count(values, 2);
                return Round(values[0], (int)Math.Truncate(values[1]));
            case "IF":
                Count(values, 3);
                return values[0] != 0 ? values[1] : values[2];
            case "AND":
                Count(values, 2);
                return values[0] != 0 && values[1] != 0 ? 1 : 0;
            case "OR":
                Count(values, 2);
                return values[0] != 0 || values[1] != 0 ? 1 : 0;
            case "SUM":
                return AtLeastOne(values).Sum();
            case "PRODUCT":
                return AtLeastOne(values).Aggregate(1.0, (a, b) => a * b);
            case "MIN":
                return AtLeastOne(values).Min();
            case "MAX":
                return AtLeastOne(values).Max();
            case "AVERAGE":
                return AtLeastOne(values).Average();
            case "COUNT":
                return AtLeastOne(values).Count;
            default:
                throw new CannotComputeException();
        }
    }

    //A name on its own: TRUE, FALSE or a bookmark that holds a number.
    private double Name(string name)
    {
        string upper = name.ToUpperInvariant();
        if (upper == "TRUE")
        {
            return 1;
        }
        if (upper == "FALSE")
        {
            return 0;
        }
        //A cell reference (A1) is not worked out here.
        if (IsCellReference(upper))
        {
            throw new CannotComputeException();
        }
        BookmarkStart bookmark = FindBookmark(name);
        if (bookmark == null)
        {
            throw new FormulaErrorException("!Undefined Bookmark, " + upper);
        }
        return BookmarkNumber(new BookmarkWrapper(bookmark).Range.Text.TrimEnd('\r', '\a').Trim());
    }

    #endregion

    #region Helpers

    private double BookmarkNumber(string text)
    {
        //Text without digits counts as 0, as in Word. A number with more than that around it, or with digit grouping, is left alone.
        if (!text.Any(char.IsDigit))
        {
            return 0;
        }
        string decimalSeparator = _numbers.NumberDecimalSeparator;
        if (!double.TryParse(text, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, _numbers, out double value))
        {
            throw new CannotComputeException();
        }
        int at = text.IndexOf(decimalSeparator, StringComparison.Ordinal);
        _decimals = Math.Max(_decimals, at < 0 ? 0 : text.Length - at - decimalSeparator.Length);
        return value;
    }

    private BookmarkStart FindBookmark(string name)
    {
        return _root.Descendants<BookmarkStart>().FirstOrDefault(b => string.Equals(b.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsCellReference(string upper)
    {
        int letters = 0;
        while (letters < upper.Length && upper[letters] >= 'A' && upper[letters] <= 'Z')
        {
            letters++;
        }
        return letters > 0 && letters <= 2 && letters < upper.Length && upper.Skip(letters).All(char.IsDigit);
    }

    //Rounds half away from zero, as Word does, to a number of decimals that can be negative (ROUND(1234, -2) is 1200).
    private static double Round(double value, int decimals)
    {
        if (decimals >= 0)
        {
            return Math.Round(value, Math.Min(decimals, 15), MidpointRounding.AwayFromZero);
        }
        double scale = Math.Pow(10, -decimals);
        return Math.Round(value / scale, MidpointRounding.AwayFromZero) * scale;
    }

    private static double Single(List<double> values)
    {
        Count(values, 1);
        return values[0];
    }

    private static int Count(List<double> values, int expected)
    {
        if (values.Count != expected)
        {
            throw new CannotComputeException();
        }
        return expected;
    }

    private static List<double> AtLeastOne(List<double> values)
    {
        if (values.Count == 0)
        {
            throw new CannotComputeException();
        }
        return values;
    }

    private bool AtEnd()
    {
        return _position >= _text.Length;
    }

    private void SkipSpaces()
    {
        while (!AtEnd() && char.IsWhiteSpace(_text[_position]))
        {
            _position++;
        }
    }

    private bool Accept(char c)
    {
        if (!AtEnd() && _text[_position] == c)
        {
            _position++;
            return true;
        }
        return false;
    }

    private void Expect(char c)
    {
        SkipSpaces();
        if (AtEnd())
        {
            throw new FormulaErrorException(TableFormulas.UnexpectedEnd);
        }
        if (!Accept(c))
        {
            throw new CannotComputeException();
        }
    }

    #endregion
}
