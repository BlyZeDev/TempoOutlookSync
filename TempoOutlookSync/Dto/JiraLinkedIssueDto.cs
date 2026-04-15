namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraLinkedIssueDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("key")]
    public required string Key { get; init; }
    [JsonPropertyName("fields")]
    public required JiraLinkedIssueFieldsDto Fields { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraLinkedIssueDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraLinkedIssueDtoJsonContext : JsonSerializerContext;