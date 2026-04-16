using Microsoft.VisualBasic;
using System.Net;
using System.Text;

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

    public string BuildHtmlBody(string? versionInfo)
    {
        var sb = new StringBuilder();

        sb.AppendLine("""
            <div style="
                font-family:'Segoe UI', Calibri, sans-serif;
                color:#111111;
                font-size:14pt;
                line-height:22pt;
                mso-line-height-rule:exactly;
            ">
            """);

        sb.AppendLine("""
            <div style="color:#666666; font-size:11pt; font-style:italic;">
                Auto-imported from Jira Tempo
            </div>
            """);

        if (!string.IsNullOrWhiteSpace(Summary))
        {
            SetSpace(20);

            sb.AppendLine($"""
                <div style="font-size:20pt; font-weight:600;">
                    {WebUtility.HtmlEncode(Summary)}
                </div>
                """);
        }

        if (!string.IsNullOrWhiteSpace(Subject))
        {
            SetSpace(20);

            sb.AppendLine($"""
                <div style="font-size:15pt;">
                    {WebUtility.HtmlEncode(Subject)}
                </div>
                """);
        }

        if (!string.IsNullOrWhiteSpace(Url))
        {
            SetSpace(10);

            var url = WebUtility.HtmlEncode(Url);

            sb.AppendLine($"""
                <div>
                    <a href="{url}" style="color:#9B59B6; font-size:13pt; text-decoration:underline;">
                        {url}
                    </a>
                </div>
                """);
        }

        if (LinkedIssues.Count > 0)
        {
            SetSpace(30);

            sb.AppendLine($"""
                <div>
                    <div style="font-size:15pt; color:#2C3E50; font-weight:600;">
                        📎 Linked issues ({LinkedIssues.Count})
                    </div>
                """);

            foreach (var issue in LinkedIssues)
            {
                var relation = WebUtility.HtmlEncode(issue.RelationToBaseIssue);
                var summary = WebUtility.HtmlEncode(issue.LinkedIssue.Summary);
                var url = WebUtility.HtmlEncode(issue.LinkedIssue.Permalink);

                sb.AppendLine("<div>");

                if (!string.IsNullOrWhiteSpace(relation))
                {
                    sb.AppendLine($"""
                        <div style="font-size:11pt; color:#7F8C8D; font-style:italic;">
                            {relation}
                        </div>
                        """);
                }

                if (!string.IsNullOrWhiteSpace(summary))
                {
                    sb.AppendLine($"""
                        <div style="font-size:12pt;">
                            {summary}
                        </div>
                        """);
                }

                sb.AppendLine($"""
                        <div>
                            <a href="{url}" style="color:#9B59B6; font-size:12pt; text-decoration:underline;">
                                {url}
                            </a>
                        </div>
                    </div>
                    """);
            }

            sb.AppendLine("</div>");
        }

        if (!string.IsNullOrWhiteSpace(PlannedBy))
        {
            SetSpace(30);

            var avatarCell = "";

            if (!string.IsNullOrWhiteSpace(PlannedByAvatarUrl))
            {
                avatarCell = $"""
                    <td style="vertical-align:middle;">
                        <img src="{WebUtility.HtmlEncode(PlannedByAvatarUrl)}"
                             width="32" height="32"
                             style="display:block;" />
                    </td>
                    <td style="width:6px;"></td>
                    """;
            }

            sb.AppendLine($"""
                <table role="presentation" cellpadding="0" cellspacing="0">
                    <tr>
                        {avatarCell}
                        <td style="vertical-align:middle;">
                            Planned by {WebUtility.HtmlEncode(PlannedBy)}
                        </td>
                    </tr>
                </table>
                """);
        }

        if (!string.IsNullOrWhiteSpace(versionInfo))
        {
            SetSpace(10);

            sb.AppendLine($"""
                <div style="font-size:11pt; font-style:italic; color:#666666;">
                    {nameof(TempoOutlookSync)} Version {WebUtility.HtmlEncode(versionInfo)}
                </div>
                """);
        }

        sb.AppendLine("</div>");

        return sb.ToString();

        void SetSpace(int spacingPx) => sb.AppendLine($"""<div style="height:{spacingPx}px; line-height:{spacingPx}px; font-size:{spacingPx}px;">&nbsp;</div>""");
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