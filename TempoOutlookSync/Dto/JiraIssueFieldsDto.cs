namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueFieldsDto
{
    [JsonPropertyName("summary")]
    public string? Summary { get; init; }
    [JsonPropertyName("project")]
    public JiraProjectDto? Project { get; init; }
    [JsonPropertyName("updated")]
    public string? Updated { get; init; }
    [JsonPropertyName("created")]
    public required string Created { get; init; }
    [JsonPropertyName("issuelinks")]
    public List<>? IssueLinks { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueFieldsDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueFieldsDtoJsonContext : JsonSerializerContext;