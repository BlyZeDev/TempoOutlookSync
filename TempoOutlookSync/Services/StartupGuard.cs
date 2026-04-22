namespace TempoOutlookSync.Services;

public sealed class StartupGuard : IDisposable
{
    private const int MutexTimeoutMs = 5000;

    private readonly Mutex _mutex;

    private bool hasHandle;

    public StartupGuard()
    {
        _mutex = new Mutex(false, TempoOutlookSyncContext.ApplicationMutexId);
    }

    public bool WaitForAccess()
    {
        try
        {
            hasHandle = _mutex.WaitOne(MutexTimeoutMs, false);
        }
        catch (AbandonedMutexException)
        {
            hasHandle = true;
        }

        return hasHandle;
    }

    public void Dispose()
    {
        try
        {
            if (hasHandle) _mutex.ReleaseMutex();
            hasHandle = false;
        }
        catch (Exception) { }
        finally
        {
            _mutex.Dispose();
            GC.SuppressFinalize(this);
        }
    }

    ~StartupGuard() => Dispose();
}