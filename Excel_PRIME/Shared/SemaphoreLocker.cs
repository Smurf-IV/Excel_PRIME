using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME.Shared;

/// <summary>
/// Borrowed from here https://stackoverflow.com/a/50139704
/// </summary>
internal class SemaphoreLocker
{
    private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

    public async Task LockAsync(Func<Task> worker)
    {
        bool isTaken = false;
        try
        {
            do
            {
                try
                {
                }
                finally
                {
                    isTaken = await _semaphore.WaitAsync(TimeSpan.FromSeconds(1)).ConfigureAwait(false);
                }
            }
            while (!isTaken);
            await worker().ConfigureAwait(false);
        }
        finally
        {
            if (isTaken)
            {
                _semaphore.Release();
            }
        }
    }

    // overloading variant for non-void methods with return type (generic T)
    public async Task<T> LockAsync<T>(Func<Task<T>> worker)
    {
        bool isTaken = false;
        try
        {
            do
            {
                try
                {
                }
                finally
                {
                    isTaken = await _semaphore.WaitAsync(TimeSpan.FromSeconds(1)).ConfigureAwait(false);
                }
            }
            while (!isTaken);
            return await worker().ConfigureAwait(false);
        }
        finally
        {
            if (isTaken)
            {
                _semaphore.Release();
            }
        }
    }
}