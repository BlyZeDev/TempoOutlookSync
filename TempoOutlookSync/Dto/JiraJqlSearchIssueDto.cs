namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed class JiraJqlSearchIssueDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraJqlSearchIssueDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraJqlSearchIssueDtoJsonContext : JsonSerializerContext;