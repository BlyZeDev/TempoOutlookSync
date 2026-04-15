using TempoOutlookSync.Dto;

namespace TempoOutlookSync.Models;

public sealed record JiraLinkedIssue
{
    public string Id { get; }
    public string Key { get; }
    public string Permalink { get; }
    public string? Summary { get; }

    public JiraLinkedIssue(JiraLinkedIssueDto dto, string baseUrl)
    {
        Id = dto.Id;
        Key = dto.Key;
        Permalink = $"{baseUrl}{Key}";
        Summary = dto.Fields.Summary;
    }
}