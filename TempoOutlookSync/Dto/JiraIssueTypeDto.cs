namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueTypeDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("name")]
    public required string Name { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueTypeDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueTypeDtoJsonContext : JsonSerializerContext;