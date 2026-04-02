namespace TempoOutlookSync.Models;

public sealed record OutlookAppointmentCreationInfo
{
    public TempoPlannerEntry TempoEntry { get; }

    public string Subject { get; }
    public string Summary { get; }
    public string? PlannedBy { get; }
    public string? Url { get; }
    public DateTime? LastUpdated { get; }
    public OutlookCategory? Category { get; }

    public OutlookAppointmentCreationInfo(TempoPlannerEntry tempoEntry, JiraUser? jiraUser)
    {
        TempoEntry = tempoEntry;
        Subject = NullIfWhiteSpace(TempoEntry.Description) ?? $"Tempo Id #{TempoEntry.Id}";
        Summary = Subject;
        PlannedBy = GetPlannedBy(jiraUser);
        Url = null;
        LastUpdated = null;
        Category = null;
    }

    public OutlookAppointmentCreationInfo(TempoPlannerEntry tempoEntry, JiraIssue jiraIssue, JiraUser? jiraUser, OutlookCategory? category)
    {
        TempoEntry = tempoEntry;
        Summary = NullIfWhiteSpace(jiraIssue.Summary) ?? NullIfWhiteSpace(jiraIssue.ProjectName) ?? jiraIssue.Key;
        Subject = NullIfWhiteSpace(tempoEntry.Description) ?? Summary;
        PlannedBy = GetPlannedBy(jiraUser);
        Url = jiraIssue.Permalink;
        LastUpdated = jiraIssue.LastUpdated;
        Category = category;
    }

    public OutlookAppointmentCreationInfo(TempoPlannerEntry tempoEntry, JiraProject jiraProject, JiraUser? jiraUser, OutlookCategory? category)
    {
        TempoEntry = tempoEntry;
        Summary = NullIfWhiteSpace(jiraProject.Name) ?? jiraProject.Key;
        Subject = NullIfWhiteSpace(tempoEntry.Description) ?? Summary;
        PlannedBy = GetPlannedBy(jiraUser);
        Url = jiraProject.Permalink;
        LastUpdated = null;
        Category = category;
    }

    private static string? NullIfWhiteSpace(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

    private static string? GetPlannedBy(JiraUser? jiraUser)
    {
        if (jiraUser is null) return null;

        var result = (jiraUser.DisplayName, jiraUser.EmailAddress) switch
        {
            (not null, not null) => $"{jiraUser.DisplayName} • {jiraUser.EmailAddress}",
            (not null, null) => jiraUser.DisplayName,
            (null, not null) => jiraUser.EmailAddress,
            _ => null
        };

        return NullIfWhiteSpace(result);
    }
}