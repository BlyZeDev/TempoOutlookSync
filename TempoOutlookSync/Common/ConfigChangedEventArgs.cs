namespace TempoOutlookSync.Common;

public sealed class ConfigChangedEventArgs : EventArgs
{
    public required Configuration OldConfig { get; init; }
    public required Configuration NewConfig { get; init; }
}