namespace TempoOutlookSync.Common;

using Microsoft.Win32.TaskScheduler;

public sealed record ScheduledTaskInfo
{
    public required bool IsActive { get; init; }
    public required TaskState State { get; init; }
    public required DateTime NextRunTime { get; init; }
}