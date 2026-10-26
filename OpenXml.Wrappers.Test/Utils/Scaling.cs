using System;
using System.Diagnostics;
using System.Globalization;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Cost checks that compare two document sizes instead of timing against a fixed limit, so the result does not depend on the speed of the machine.
/// Four times the size takes about four times as long when the work is linear, about sixteen times as long when each step walks the whole
/// document, and about as long when the work only touches the part it needs. Timings on a shared machine are noisy, so the limits are
/// generous: they catch a cost that grows far faster than it does today, not small changes.
/// </summary>
internal static class Scaling
{
    // Below this the work is quick on any machine and the ratio is only noise.
    private static readonly TimeSpan Negligible = TimeSpan.FromMilliseconds(250);

    private const int Repeats = 3;

    public static TimeSpan Time(Action action)
    {
        Stopwatch watch = Stopwatch.StartNew();
        action();
        watch.Stop();
        return watch.Elapsed;
    }

    /// <summary>
    /// Runs <paramref name="measure"/> (which builds a document of the given size, times only the operation and returns the time) at a small
    /// and a large size, takes the best of a few runs of each, and fails if the large one took more than <paramref name="maxRatio"/> times as long.
    /// </summary>
    public static void AssertGrowth(Func<int, TimeSpan> measure, double maxRatio, int small = 1500, int large = 6000)
    {
        measure(small / 4);
        TimeSpan fast = Best(measure, small);
        TimeSpan slow = Best(measure, large);
        if (slow < Negligible)
        {
            return;
        }

        double ratio = slow.TotalMilliseconds / Math.Max(fast.TotalMilliseconds, 1);
        Assert.True(ratio <= maxRatio, string.Format(CultureInfo.InvariantCulture,
            "{0}x the size took {1:0.0}x as long ({2} then {3}); at most {4}x is allowed", large / small, ratio, fast, slow, maxRatio));
    }

    private static TimeSpan Best(Func<int, TimeSpan> measure, int size)
    {
        TimeSpan best = TimeSpan.MaxValue;
        for (int i = 0; i < Repeats; i++)
        {
            TimeSpan took = measure(size);
            if (took < best)
            {
                best = took;
            }
        }
        return best;
    }
}
