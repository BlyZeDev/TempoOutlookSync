namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueFieldsDto
{
    [JsonPropertyName("summary")]
    public string? Summary { get; init; }
    [JsonPropertyName("issuetype")]
    public JiraIssueTypeDto? IssueType { get; init; }
    [JsonPropertyName("project")]
    public JiraIssueProjectDto? Project { get; init; }
    [JsonPropertyName("status")]
    public JiraIssueStatusDto? Status { get; init; }
    [JsonPropertyName("updated")]
    public string? Updated { get; init; }
    [JsonPropertyName("created")]
    public required string Created { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueFieldsDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueFieldsDtoJsonContext : JsonSerializerContext;