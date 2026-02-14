namespace TempoOutlookSync.Common;

using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;

public static class Util
{
    private const string StartupRegistryKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

    public static string GetFileNameTimestamp() => $"{DateTime.UtcNow:yyyyMMddHHmmssff}Z";

    public static string? FormatTime(in TimeSpan timeSpan)
    {
        if (timeSpan == TimeSpan.Zero) return null;

        var (value, unit) = timeSpan switch
        {
            var _ when timeSpan.TotalDays >= 1 => (timeSpan.TotalDays, "day"),
            var _ when timeSpan.TotalHours >= 1 => (timeSpan.TotalHours, "hour"),
            var _ when timeSpan.TotalMinutes >= 1 => (timeSpan.TotalMinutes, "minute"),
            var _ when timeSpan.TotalSeconds >= 1 => (timeSpan.TotalSeconds, "second"),
            var _ when timeSpan.TotalMilliseconds >= 1 => (timeSpan.TotalMilliseconds, "millisecond"),
            var _ when timeSpan.TotalMicroseconds >= 1 => (timeSpan.TotalMicroseconds, "microsecond"),
            _ => (timeSpan.TotalNanoseconds, "nanosecond")
        };

        if (value != 1) unit += 's';

        return $"{(int)Math.Ceiling(value)} {unit}";
    }

    public static void ShellOpen(string fileName)
    {
        using (var process = new Process())
        {
            process.StartInfo = new ProcessStartInfo
            {
                UseShellExecute = true,
                FileName = fileName
            };
            process.Start();
        }
    }

    public static bool IsInStartup(string name, string path)
    {
        using (var key = Registry.CurrentUser.OpenSubKey(StartupRegistryKey))
        {
            if (key is null) return false;

            var value = key.GetValue(name)?.ToString();
            return value is not null && path.Equals(value.Trim('\"'), StringComparison.OrdinalIgnoreCase);
        }
    }

    public static bool AddToStartup(string name, string path)
    {
        using (var key = Registry.CurrentUser.OpenSubKey(StartupRegistryKey, true))
        {
            if (key is null) return false;

            key.SetValue(name, $"\"{path}\"");
            return true;
        }
    }

    public static bool RemoveFromStartup(string name)
    {
        using (var key = Registry.CurrentUser.OpenSubKey(StartupRegistryKey, true))
        {
            if (key is null) return false;

            key.DeleteValue(name);
            return true;
        }
    }
}