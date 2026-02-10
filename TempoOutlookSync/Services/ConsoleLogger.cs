namespace TempoOutlookSync.Services;

using System;
using System.Text;
using TempoOutlookSync.Common;

public sealed class ConsoleLogger : ILoggerTarget
{
    public void LogMessage(LogLevel logLevel, string text, Exception? exception, CallerInfo? callerInfo)
    {
        var builder = new StringBuilder();
        builder.Append($"{DateTime.Now:dd.MM.yyyy HH:mm:ss.ffff} | ");
        builder.Append(logLevel);

        if (callerInfo is not null) builder.Append($" | {callerInfo}");

        builder.AppendLine($" | {text}");

        if (exception is not null)
        {
            builder.AppendLine("Exception");
            builder.AppendLine(exception.ToString());
        }

        Console.Write(builder.ToString());
    }
}