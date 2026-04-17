namespace TempoOutlookSync.Common;

[Flags]
public enum SyncBlocker
{
    None = 0,
    Manual = 1 << 0,
    Network = 1 << 1,
    Session = 1 << 2,
    Power = 1 << 3,
}