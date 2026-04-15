namespace TempoOutlookSync.Models;

public sealed record OutlookAppointmentCreationInfo
{
    public TempoPlannerEntry TempoEntry { get; }

    public string Subject { get; }
    public string Summary { get; }
    public string? PlannedBy { get; }
    public string? PlannedByAvatarUrl { get; }
    public string? Url { get; }
    public DateTime? LastUpdated { get; }
    public OutlookCategory? Category { get; }
    public IReadOnlyCollection<JiraLink> LinkedIssues { get; }

    public OutlookAppointmentCreationInfo(TempoPlannerEntry tempoEntry, JiraUser? jiraUser)
    {
        TempoEntry = tempoEntry;
        Subject = NullIfWhiteSpace(TempoEntry.Description) ?? $"Tempo Id #{TempoEntry.Id}";
        Summary = Subject;
        PlannedBy = GetPlannedBy(jiraUser);
        PlannedByAvatarUrl = jiraUser?.AvatarUrl;
        Url = null;
        LastUpdated = null;
        Category = null;
        LinkedIssues = [];
    }

    public OutlookAppointmentCreationInfo(TempoPlannerEntry tempoEntry, JiraIssue jiraIssue, JiraUser? jiraUser, OutlookCategory? category) : this(tempoEntry, jiraUser)
    {
        Summary = NullIfWhiteSpace(jiraIssue.Summary) ?? NullIfWhiteSpace(jiraIssue.ProjectName) ?? jiraIssue.Key;
        Subject = NullIfWhiteSpace(tempoEntry.Description) ?? Summary;
        Url = jiraIssue.Permalink;
        LastUpdated = jiraIssue.LastUpdated;
        Category = category;
        LinkedIssues = jiraIssue.LinkedIssues.ToArray();
    }

    public OutlookAppointmentCreationInfo(TempoPlannerEntry tempoEntry, JiraProject jiraProject, JiraUser? jiraUser, OutlookCategory? category) : this(tempoEntry, jiraUser)
    {
        Summary = NullIfWhiteSpace(jiraProject.Name) ?? jiraProject.Key;
        Subject = NullIfWhiteSpace(tempoEntry.Description) ?? Summary;
        Url = jiraProject.Permalink;
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