using System.Text;

namespace ExcelPRIME.Implementation;

/// <summary>
/// Per-thread pooled StringBuilder using [ThreadStatic] to avoid locks.
/// - Single builder is kept per-thread; Rent hands it out and clears on return.
/// - Builders that grow beyond MaxPooledCapacity are discarded on Return to avoid retaining large buffers.
/// - This approach is lock-free and low-overhead for highly parallel workloads.
/// </summary>
internal static class ThreadStringBuilderPool
{
    [ThreadStatic]
    private static StringBuilder? t_builder;

    // Use a smaller initial capacity to avoid allocating large char[] for typical small cell text.
    private const int InitialCapacity = 128;
    // Keep a modest max pooled capacity to avoid retaining large buffers across requests.
    private const int MaxPooledCapacity = 1024;

    public static StringBuilder Rent()
    {
        StringBuilder? sb = t_builder;
        if (sb == null)
        {
            return new StringBuilder(InitialCapacity);
        }
        t_builder = null;
        return sb;
    }

    public static void Return(StringBuilder? sb)
    {
        if (sb == null)
        {
            return;
        }

        if (sb.Capacity > MaxPooledCapacity)
        {
            return;
        }

        sb.Clear();
        // Replace any existing thread-local builder (drop the previous one).
        t_builder = sb;
    }
}