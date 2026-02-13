namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("key")]
    public required string Key { get; init; }
    [JsonPropertyName("fields")]
    public required JiraIssueFieldsDto Fields { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueDtoJsonContext : JsonSerializerContext;