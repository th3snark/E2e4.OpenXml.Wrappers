using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Wrappers.Helpers;

public static class ConversionHelper
{
    public static int ToPct(int width, int tableWidth)
    {
        return (int)Math.Round(((double)width / tableWidth) * 5000.0);
    }

    public static int CentimeterToDxa(float cm)
    {
        //2.54 cm is 1 inch.
        //72 points per inch.
        //20.0 dxa per point.
        return (int)Math.Round(cm * 1440.0 / 2.54, 0);
    }
}
