namespace TempoOutlookSync.Models;

public sealed record OutlookAppointmentRef
{
    public required int TempoId { get; init; }
    public required string EntryId { get; init; }
    public required DateTime Start { get; init; }
    public required DateTime End { get; init; }
    public required DateTime? TempoUpdated { get; init; }
    public required DateTime? JiraUpdated { get; init; }

    public bool Equals(OutlookAppointmentRef? other) => other is not null && EntryId == other.EntryId;

    public override int GetHashCode() => EntryId.GetHashCode(StringComparison.Ordinal);
}