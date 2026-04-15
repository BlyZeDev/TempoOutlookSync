namespace TempoOutlookSync.Models;

using System.Globalization;
using TempoOutlookSync.Dto;

public sealed record JiraIssue
{
    public string Id { get; }
    public string Key { get; }
    public string Permalink { get; }
    public string? Summary { get; }
    public string? ProjectName { get; }
    public DateTime LastUpdated { get; }
    public IEnumerable<JiraLink> LinkedIssues { get; }

    public JiraIssue(JiraIssueDto dto, string baseUrl)
    {
        Id = dto.Id;
        Key = dto.Key;
        Permalink = $"{baseUrl}{Key}";
        Summary = dto.Fields.Summary;
        ProjectName = dto.Fields.Project?.Name;
        LastUpdated = DateTimeOffset.ParseExact(dto.Fields.Updated ?? dto.Fields.Created, "yyyy-MM-ddTHH:mm:ss.FFFFFFFzz00", CultureInfo.InvariantCulture).UtcDateTime;
        LinkedIssues = dto.Fields.IssueLinks?.Select(x => new JiraLink(x, baseUrl)) ?? [];
    }
}