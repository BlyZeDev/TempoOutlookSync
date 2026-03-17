namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraJqlSearchPayloadDto
{
    [JsonPropertyName("issues")]
    public List<JiraJqlSearchIssueDto>? Issues { get; init; }
    [JsonPropertyName("nextPageToken")]
    public string? NextPageToken { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraJqlSearchPayloadDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraJqlSearchPayloadDtoJsonContext : JsonSerializerContext;