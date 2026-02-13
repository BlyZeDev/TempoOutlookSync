namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueProjectCategoryDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("name")]
    public required string Name { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueProjectCategoryDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueProjectCategoryDtoJsonContext : JsonSerializerContext;