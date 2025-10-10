using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
#pragma warning disable CA1024

namespace Excel_PRIME.Shared;

#pragma warning disable CA1711 // Do not end Names with Dictionary

/// <summary>
/// Taken and then modified from
/// https://stackoverflow.com/questions/255341/getting-multiple-keys-of-specified-value-of-a-generic-dictionary/255638#255638
/// </summary>
/// <typeparam name="TFirst"></typeparam>
/// <typeparam name="TSecond"></typeparam>
public class BiDictionary<TFirst, TSecond> where TFirst : notnull where TSecond : notnull
{
    private readonly IDictionary<TFirst, TSecond> _firstToSecond = new Dictionary<TFirst, TSecond>();
    private readonly IDictionary<TSecond, TFirst> _secondToFirst = new Dictionary<TSecond, TFirst>();

    public BiDictionary(IDictionary<TFirst, TSecond> dictionary)
    {
        ArgumentNullException.ThrowIfNull(dictionary);

        foreach (KeyValuePair<TFirst, TSecond> keyValuePair in dictionary)
        {
            Add(keyValuePair.Key, keyValuePair.Value);
        }
    }

    public void Add(TFirst first, TSecond second)
    {
        _firstToSecond.Add(first, second);
        _secondToFirst.Add(second, first);
    }

    // Note potential ambiguity using indexers (e.g. mapping from int to int)
    // Hence the methods as well...
    public TSecond this[TFirst first] => GetByFirst(first)!;

    public TFirst this[TSecond second] => GetBySecond(second)!;

    public TSecond? GetByFirst(TFirst first)
    {
        _firstToSecond.TryGetValue(first, out TSecond? second);
        return second;
    }

    public TFirst? GetBySecond(TSecond second)
    {
        _secondToFirst.TryGetValue(second, out TFirst? first);
        return first;
    }

    public ICollection<TFirst> GetAllFirsts() => _firstToSecond.Keys;

    public IReadOnlyDictionary<TFirst, TSecond> FirstToSecond => new ReadOnlyDictionary<TFirst, TSecond>(_firstToSecond);
    public IReadOnlyDictionary<TSecond, TFirst> SecondToFirst => new ReadOnlyDictionary<TSecond, TFirst>(_secondToFirst);
}
