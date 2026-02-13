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
    public string? IssueType { get; }
    public string? ProjectName { get; }
    public string? ProjectCategory { get; }
    public JiraStatusCategory StatusCategory { get; }
    public JiraStatus Status { get; }
    public DateTime LastUpdated { get; }

    public JiraIssue(JiraIssueDto dto, string baseUrl)
    {
        Id = dto.Id;
        Key = dto.Key;
        Permalink = $"{baseUrl}{Key}";
        Summary = dto.Fields.Summary;
        IssueType = dto.Fields.IssueType?.Name;
        ProjectName = dto.Fields.Project?.Name;
        ProjectCategory = dto.Fields.Project?.Category?.Name;
        StatusCategory = ParseStatusCategory(dto.Fields.Status?.Category?.Name);
        Status = ParseStatus(dto.Fields.Status?.Name);
        LastUpdated = DateTimeOffset.ParseExact(dto.Fields.Updated ?? dto.Fields.Created, "yyyy-MM-ddTHH:mm:ss.FFFFFFFzz00", CultureInfo.InvariantCulture).UtcDateTime;
    }

    public override string ToString() => JsonSerializer.Serialize(this, _options);

    private static JiraStatusCategory ParseStatusCategory(string? categoryName)
    {
        return categoryName switch
        {
            null => JiraStatusCategory.Undefined,
            var name when name.Equals("new", StringComparison.OrdinalIgnoreCase) => JiraStatusCategory.New,
            var name when name.Equals("indeterminate", StringComparison.OrdinalIgnoreCase) => JiraStatusCategory.Indeterminate,
            var name when name.Equals("done", StringComparison.OrdinalIgnoreCase) => JiraStatusCategory.Done,
            _ => JiraStatusCategory.Undefined
        };
    }

    private static JiraStatus ParseStatus(string? statusName)
    {
        return statusName switch
        {
            null => JiraStatus.Unknown,
            var name when name.Equals("Warten auf Kunde", StringComparison.OrdinalIgnoreCase) => JiraStatus.WaitingForCustomer,
            var name when name.Equals("In Arbeit", StringComparison.OrdinalIgnoreCase) => JiraStatus.InProgess,
            var name when name.Equals("Aufgabe Kunde", StringComparison.OrdinalIgnoreCase) => JiraStatus.CustomerAssignment,
            var name when name.Equals("Warten auf 3rd Level", StringComparison.OrdinalIgnoreCase) => JiraStatus.WaitingFor3rdLevel,
            var name when name.Equals("Aufgabe Edoc", StringComparison.OrdinalIgnoreCase) => JiraStatus.EdocAssignment,
            _ => JiraStatus.Unknown
        };
    }
}