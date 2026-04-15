namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraLinkedIssueFieldsDto
{
    [JsonPropertyName("summary")]
    public string? Summary { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraLinkedIssueFieldsDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraLinkedIssueFieldsDtoJsonContext : JsonSerializerContext;