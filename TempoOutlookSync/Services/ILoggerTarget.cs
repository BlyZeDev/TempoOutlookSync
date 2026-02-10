namespace TempoOutlookSync.Services;

using TempoOutlookSync.Common;

public interface ILoggerTarget
{
    public void LogMessage(LogLevel logLevel, string text, Exception? exception, CallerInfo? callerInfo);
}