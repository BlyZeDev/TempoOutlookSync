namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueStatusDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("name")]
    public required string Name { get; init; }
    [JsonPropertyName("statusCategory")]
    public JiraIssueStatusCategoryDto? Category { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueStatusDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueStatusDtoJsonContext : JsonSerializerContext;