using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Wrappers.Helpers;

internal static class ConversionHelper
{
    internal static int ToPct(int width, int tableWidth)
    {
        return (int)Math.Round(((double)width / tableWidth) * 5000.0);
    }

    /// <summary>
    /// Points (Word's unit, 72 to the inch) as twips, the twentieths of a point that Open XML stores.
    /// </summary>
    internal static int PointsToTwips(double points, string parameterName)
    {
        if (Double.IsNaN(points) || points < 0 || points > 31680)
        {
            throw new ArgumentOutOfRangeException(parameterName, "The value must be between 0 and 31680 points.");
        }
        return (int)Math.Round(points * 20, MidpointRounding.AwayFromZero);
    }

    internal static double TwipsToPoints(int twips)
    {
        return twips / 20.0;
    }

    internal static int CentimeterToDxa(float cm)
    {
        //2.54 cm is 1 inch.
        //72 points per inch.
        //20.0 dxa per point.
        return (int)Math.Round(cm * 1440.0 / 2.54, 0);
    }
}
