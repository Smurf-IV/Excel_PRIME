namespace ExcelPRIME.Implementation;

internal class MagicDictionary<TKey, TValue> : SortedList<TKey, TValue> where TKey : notnull
{
    public TValue GetOrAdd(TKey key, Func<TKey, TValue> valueFactory)
    {
        if (!TryGetValue(key, out TValue? value))
        {
            value = valueFactory(key);
            this[key] = value;  //Overwrite just in case of threading issues, but it should be the same value if another thread added it in the meantime.
        }
        return value!;
    }
}