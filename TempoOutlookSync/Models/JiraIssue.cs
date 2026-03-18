namespace TempoOutlookSync.Models;

using System.Globalization;
using System.Text.Json;
using TempoOutlookSync.Dto;

public sealed record JiraIssue
{
    private static readonly JsonSerializerOptions _options = new JsonSerializerOptions()
    {
        WriteIndented = true
    };

    public string Id { get; }
    public string Key { get; }
    public string Permalink { get; }
    public string? Summary { get; }
    public string? ProjectName { get; }
    public DateTime LastUpdated { get; }

    public JiraIssue(JiraIssueDto dto, string baseUrl)
    {
        Id = dto.Id;
        Key = dto.Key;
        Permalink = $"{baseUrl}{Key}";
        Summary = dto.Fields.Summary;
        ProjectName = dto.Fields.Project?.Name;
        LastUpdated = DateTimeOffset.ParseExact(dto.Fields.Updated ?? dto.Fields.Created, "yyyy-MM-ddTHH:mm:ss.FFFFFFFzz00", CultureInfo.InvariantCulture).UtcDateTime;
    }

    public override string ToString() => JsonSerializer.Serialize(this, _options);
}