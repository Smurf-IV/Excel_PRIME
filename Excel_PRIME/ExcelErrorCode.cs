namespace ExcelPRIME;

/// <summary>
/// Indicates the kind of error a formula produced.
/// </summary>
public enum ExcelErrorCode /*: byte*/
{
    /// <summary>A null reference error.</summary>
    Null = 0,

    /// <summary>A division by zero error.</summary>
    DivideByZero = 7,

    /// <summary>
    /// A value error indicating a function requires a numeric but was given a string.
    /// </summary>
    Value = 15, // 0x0000000F

    /// <summary>
    /// A reference error indicating a function references a location that doesn't exist.
    /// </summary>
    Reference = 23, // 0x00000017

    /// <summary>
    /// A name error indicating the function references an unknown operation.
    /// </summary>
    Name = 29, // 0x0000001D

    /// <summary>
    /// A number error indicating the function expected a number in a certain range.
    /// </summary>
    Number = 36, // 0x00000024

    /// <summary>
    /// An error indicating the function attempted to lookup a value that isn't available.
    /// </summary>
    NotAvailable = 42 // 0x0000002A
}