using System;
using System.Buffers;
using System.Reflection;
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

    private const int InitialCapacity = 512;
    private const int MaxPooledCapacity = 2048;
    private static readonly ArrayPool<char> s_charPool = ArrayPool<char>.Shared;
    private static readonly PropertyInfo? s_poolBufferProperty = typeof(StringBuilder).GetProperty("_poolBuffer", BindingFlags.NonPublic | BindingFlags.Instance);
    private static readonly FieldInfo? s_chunkCharsField = typeof(StringBuilder).GetField("m_ChunkChars", BindingFlags.NonPublic | BindingFlags.Instance);

    public static StringBuilder Rent()
    {
        StringBuilder? sb = t_builder;
        if (sb == null)
        {
            char[] buffer = s_charPool.Rent(InitialCapacity);
            sb = new StringBuilder(InitialCapacity);
            s_poolBufferProperty?.SetValue(sb, buffer);
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
            // If buffer grew too large, do not keep it for the thread to avoid retaining large memory.
            // Try to return the buffer to the pool if possible.
            ReturnBufferToPool(sb);
            return;
        }

        sb.Length = 0;
        // Replace any existing thread-local builder (drop the previous one).
        t_builder = sb;
    }

    private static void ReturnBufferToPool(StringBuilder sb)
    {
        if (s_chunkCharsField != null)
        {
            if (s_chunkCharsField.GetValue(sb) is char[] buffer)
            {
                s_charPool.Return(buffer);
            }
        }
    }
}