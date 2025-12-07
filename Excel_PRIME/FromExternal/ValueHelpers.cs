using System;

namespace ExcelPRIME.FromExternal;

/// <summary>
/// Provides utility methods for evaluating and handling raw values of various types.
/// </summary>
/// <remarks>The ValueHelpers class contains static methods that assist in determining the presence or state of
/// values, particularly when working with loosely-typed or boxed data. These helpers are useful when processing input
/// that may be represented as strings, character arrays, or other primitive types, and can help standardize checks for
/// non-empty or meaningful values across different scenarios.</remarks>
public static class ValueHelpers
{
    /// <summary>
    /// Determines whether the specified raw value is non-null and not empty, based on its type.
    /// </summary>
    /// <remarks>For strings, character arrays, and ReadOnlyMemory/<char/>, the method returns true only if the
    /// value contains at least one character. For other types, any non-null value is considered non-empty.</remarks>
    /// <param name="raw">The value to check for non-emptiness. Can be a string, character array, ReadOnlyMemory/<char/>, or any other
    /// object.</param>
    /// <returns>true if the value is non-null and, for supported types, not empty; otherwise, false.</returns>
    public static bool HasNonEmptyRawValue(object? raw) =>
        raw switch
        {
            null => false,
            string s => !string.IsNullOrEmpty(s),
            ReadOnlyMemory<char> rom => rom.Length > 0,
            char[] carr => carr.Length > 0,         // boxed ReadOnlySpan<char> is not possible; ignore
            _ => true // For primitive numeric and boolean types and other objects, treat as non-empty
        };
}