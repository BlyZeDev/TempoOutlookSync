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

    private static OutlookCategory? CreateCategory(string? issueType, string? issueKey, string? projectKey, string? projectName, string? category, JiraStatus status)
    {
        if (issueKey?.Equals("EDOCSE-136", StringComparison.OrdinalIgnoreCase) ?? false) return new OutlookCategory
        {
            Name = "Intern",
            Color = OlCategoryColor.olCategoryColorPurple
        };

        if ((issueType?.Equals("Support", StringComparison.OrdinalIgnoreCase) ?? false)
            || (projectName?.Equals("edoc Kundenportal", StringComparison.OrdinalIgnoreCase) ?? false))
        {
            return status switch
            {
                JiraStatus.WaitingForCustomer => new OutlookCategory
                {
                    Name = "Support - Warte auf Kunde",
                    Color = OlCategoryColor.olCategoryColorDarkBlue
                },
                JiraStatus.InProgess => new OutlookCategory
                {
                    Name = "Support - In Arbeit",
                    Color = OlCategoryColor.olCategoryColorTeal
                },
                _ => new OutlookCategory
                {
                    Name = "Support - Andere",
                    Color = OlCategoryColor.olCategoryColorBlue
                }
            };
        }

        if ((category?.Equals("BC", StringComparison.OrdinalIgnoreCase) ?? false)
            || (projectKey?.Equals("SP", StringComparison.OrdinalIgnoreCase) ?? false))
        {
            return status switch
            {
                JiraStatus.CustomerAssignment or JiraStatus.WaitingFor3rdLevel => new OutlookCategory
                {
                    Name = "Kundenprojekt - Aufgabe Kunde/Warte auf 3rd Level",
                    Color = OlCategoryColor.olCategoryColorOrange
                },
                JiraStatus.InProgess or JiraStatus.EdocAssignment => new OutlookCategory
                {
                    Name = "Kundenprojekt - In Arbeit/Aufgabe Edoc",
                    Color = OlCategoryColor.olCategoryColorDarkOrange
                },
                _ => new OutlookCategory
                {
                    Name = "Kundenprojekt - Andere",
                    Color = OlCategoryColor.olCategoryColorRed
                }
            };
        }

        return null;
    }
}