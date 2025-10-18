namespace ExcelPRIME.Shared;

internal static class Extensions
{
    // https://stackoverflow.com/a/6723764
    public static unsafe int IntParseUnsafe(this string value)
    {
        int result = 0;
        fixed (char* v = value)
        {
            char* str = v;
            while (*str != '\0')
            {
                result = 10 * result + (*str - 48);
                str++;
            }
        }
        return result;
    }
}
