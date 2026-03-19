namespace TempoOutlookSync.Services;

using System.Diagnostics;
using System.Text;
using TempoOutlookSync.Common;

public sealed class DebugLogger : ILoggerTarget
{
    public DebugLogger() { }

    public void LogMessage(LogLevel logLevel, string text, Exception? exception, CallerInfo? callerInfo)
    {
        var builder = new StringBuilder();
        builder.Append($"{DateTime.Now:dd.MM.yyyy HH:mm:ss.ffff} | ");
        builder.Append(logLevel.ToString());

        if (callerInfo is not null) builder.Append($" | {callerInfo}");

        builder.AppendLine($" | {text}");

        if (exception is not null)
        {
            builder.AppendLine("Exception");
            builder.AppendLine(exception.ToString());
        }

        Debug.Write(builder.ToString());
        Debug.Flush();
    }
}