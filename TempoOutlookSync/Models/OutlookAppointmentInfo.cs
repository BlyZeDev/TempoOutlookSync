namespace TempoOutlookSync.Models;

using Microsoft.Office.Interop.Outlook;

public sealed record OutlookAppointmentInfo
{
    public TempoPlannerEntry TempoEntry { get; }

    public string Subject { get; }
    public string Summary { get; }
    public string? Url { get; }
    public DateTime? LastUpdated { get; }
    public OutlookCategory? Category { get; }

    public OutlookAppointmentInfo(TempoPlannerEntry tempoEntry)
    {
        TempoEntry = tempoEntry;
        Subject = string.IsNullOrWhiteSpace(TempoEntry.Description) ? $"Tempo Id #{TempoEntry.Id}" : TempoEntry.Description;
        Summary = Subject;
        Url = null;
        LastUpdated = null;
        Category = null;
    }

    public OutlookAppointmentInfo(TempoPlannerEntry tempoEntry, JiraIssue jiraIssue)
    {
        TempoEntry = tempoEntry;
        Summary = NullIfWhiteSpace(jiraIssue.Summary) ?? NullIfWhiteSpace(jiraIssue.ProjectName) ?? jiraIssue.Key;
        Subject = NullIfWhiteSpace(tempoEntry.Description) ?? Summary;
        Url = jiraIssue.Permalink;
        LastUpdated = jiraIssue.LastUpdated;
        Category = CreateCategory(jiraIssue.IssueType, jiraIssue.Key, jiraIssue.ProjectKey, jiraIssue.ProjectName, jiraIssue.ProjectCategory, jiraIssue.Status);
    }

    public OutlookAppointmentInfo(TempoPlannerEntry tempoEntry, JiraProject jiraProject)
    {
        TempoEntry = tempoEntry;
        Summary = NullIfWhiteSpace(jiraProject.Name) ?? jiraProject.Key;
        Subject = NullIfWhiteSpace(tempoEntry.Description) ?? Summary;
        Url = jiraProject.Permalink;
        LastUpdated = null;
        Category = CreateCategory(null, null, jiraProject.Key, jiraProject.Name, jiraProject.Category, JiraStatus.Other);
    }

    private static string? NullIfWhiteSpace(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;
}