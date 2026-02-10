namespace TempoOutlookSync.Services;

public sealed class StartupGuard : IDisposable
{
    private const string Guid = "07863666-66fb-41ba-9cc1-83725487810d";
    private const string MutexId = $@"Global\{{{nameof(TempoOutlookSync)}-{Guid}}}";
    private const int MutexTimeoutMs = 5000;

    private readonly Mutex _mutex;

    private bool hasHandle;

    public StartupGuard()
    {
        _mutex = new Mutex(false, MutexId, out _);
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