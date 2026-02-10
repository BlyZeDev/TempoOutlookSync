namespace TempoOutlookSync.Common;

public static class Util
{
    private const string StartupRegistryKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

    public static string GetFileNameTimestamp() => $"{DateTime.UtcNow:yyyyMMddHHmmssff}Z";
}