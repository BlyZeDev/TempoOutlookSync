namespace TempoOutlookSync.Common;

using Microsoft.Win32;
using System.Diagnostics;

public static class Util
{
    private const string StartupRegistryKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

    public static string GetFileNameTimestamp() => $"{DateTime.UtcNow:yyyyMMddHHmmssff}Z";

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