namespace TempoOutlookSync.Models;

public sealed record OutlookAppointmentInfo
{
    public required TempoPlannerEntry TempoEntry { get; init; }
    public required JiraIssueOrProject? JiraIssue { get; init; }
}