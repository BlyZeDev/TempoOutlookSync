namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraProjectDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("key")]
    public required string Key { get; init; }
    [JsonPropertyName("name")]
    public required string Name { get; init; }
    [JsonPropertyName("projectCategory")]
    public JiraIssueProjectCategoryDto? Category { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraProjectDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraProjectDtoJsonContext : JsonSerializerContext;