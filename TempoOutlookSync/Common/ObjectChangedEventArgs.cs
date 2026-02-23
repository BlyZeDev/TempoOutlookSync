namespace TempoOutlookSync.Common;

public sealed class ObjectChangedEventArgs<T> : EventArgs
{
    public required T Old { get; init; }
    public required T New { get; init; }
}