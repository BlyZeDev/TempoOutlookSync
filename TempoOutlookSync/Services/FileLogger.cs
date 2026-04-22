namespace TempoOutlookSync.Services;

using System.Text;
using TempoOutlookSync.Common;

public sealed class FileLogger : ILoggerTarget, IDisposable
{
    private const string LoggerMutexId = $"{TempoOutlookSyncContext.ApplicationMutexId}-FileLogger";
    private const int MutexTimeoutMs = 1000;

    private readonly TempoOutlookSyncContext _context;

    private readonly Mutex _mutex;

    private bool hasHandle;

    public FileLogger(TempoOutlookSyncContext context)
    {
        _context = context;

        _mutex = new Mutex(false, LoggerMutexId);
    }

    public void LogMessage(LogLevel logLevel, string text, Exception? exception, CallerInfo? callerInfo)
    {
        var builder = new StringBuilder();
        builder.Append($"{DateTime.Now:dd.MM.yyyy HH:mm:ss.ffff} | [{(_context.IsHeadless ? "Sync" : "App")}] | ");
        builder.Append(logLevel.ToString());

        if (callerInfo is not null) builder.Append($" | {callerInfo}");

        builder.AppendLine($" | {text}");

        if (exception is not null)
        {
            builder.AppendLine("Exception");
            builder.AppendLine(exception.ToString());
        }

        try
        {
            hasHandle = _mutex.WaitOne(MutexTimeoutMs, false);
        }
        catch (AbandonedMutexException)
        {
            hasHandle = true;
        }

        if (!hasHandle) return;

        try
        {
            var path = Path.Combine(_context.LogDirectory, $"{DateTime.Now:yyyy-MM-dd}.log");
            using (var fileStream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read))
            {
                using (var writer = new StreamWriter(fileStream, Encoding.UTF8))
                {
                    writer.Write(builder.ToString());
                }
            }
        }
        finally
        {
            _mutex.ReleaseMutex();
            hasHandle = false;
        }
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

    ~FileLogger() => Dispose();
}