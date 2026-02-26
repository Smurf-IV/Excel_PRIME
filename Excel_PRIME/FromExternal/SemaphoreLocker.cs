using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME.FromExternal;

/// <summary>
/// Borrowed from here https://stackoverflow.com/a/50139704
/// </summary>
internal sealed class SemaphoreLocker : IDisposable
{
    private readonly SemaphoreSlim _semaphore = new(1, 1);
    private bool _isDisposed;

    public void Lock(Action worker)
    {
        bool isTaken = false;
        try
        {
            do
            {
                isTaken = _semaphore.Wait(TimeSpan.FromMilliseconds(250));
            } while (!isTaken);
            worker();
        }
        finally
        {
            if (isTaken)
            {
                _semaphore.Release();
            }
        }
    }

    public T Lock<T>(Func<T> worker)
    {
        bool isTaken = false;
        try
        {
            do
            {
                isTaken = _semaphore.Wait(TimeSpan.FromMilliseconds(250));
            } while (!isTaken);
            return worker();
        }
        finally
        {
            if (isTaken)
            {
                _semaphore.Release();
            }
        }
    }

    public async Task LockAsync(Func<Task> worker)
    {
        bool isTaken = false;
        try
        {
            do
            {
                isTaken = await _semaphore.WaitAsync(TimeSpan.FromMilliseconds(250)).ConfigureAwait(false);
            } while (!isTaken);
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
                isTaken = await _semaphore.WaitAsync(TimeSpan.FromMilliseconds(250)).ConfigureAwait(false);
            } while (!isTaken);
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

    // New non-allocating lock helpers
    public void Enter() => _semaphore.Wait();

    public void Exit() => _semaphore.Release();

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _semaphore.Dispose();
            }

            _isDisposed = true;
        }
    }

    public void Dispose() => Dispose(isDisposing: true);
}