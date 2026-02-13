namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueStatusCategoryDto
{
    [JsonPropertyName("id")]
    public required int Id { get; init; }
    [JsonPropertyName("name")]
    public required string Name { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueStatusCategoryDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueStatusCategoryDtoJsonContext : JsonSerializerContext;