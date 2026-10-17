namespace OpenXml.Wrappers.Samples;

/// <summary>
/// Writes every example document to a folder: <c>dotnet run --project samples/OpenXml.Wrappers.Samples [folder]</c>. The folder defaults to
/// <c>output</c> in the current directory.
/// </summary>
internal static class Program
{
    private static int Main(string[] args)
    {
        string folder = args.Length > 0 ? args[0] : "output";
        Directory.CreateDirectory(folder);

        foreach (Example example in Examples.All)
        {
            string path = Path.Combine(folder, example.FileName);
            example.Create(path);
            Console.WriteLine("{0,-12} {1}", example.Category, path);
        }

        Console.WriteLine();
        Console.WriteLine("Wrote {0} documents to {1}.", Examples.All.Count, Path.GetFullPath(folder));
        return 0;
    }
}
