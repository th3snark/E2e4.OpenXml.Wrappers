using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXml.Wrappers.Test.Utils;

public static class TestFile
{
    public const string OutputDirectory = @"TestFiles";

    public const string Document1 = @"TestFiles\Document1.docm";
    public const string Document2 = @"TestFiles\Document2.docx";
    public const string Document3 = @"TestFiles\Document3.docx";

    /// <summary>A 40 x 20 pixel PNG at 96 dpi.</summary>
    public const string Picture = @"TestFiles\Pictures\blue-40x20.png";
}
