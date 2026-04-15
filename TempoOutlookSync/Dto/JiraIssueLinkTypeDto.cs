namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueLinkTypeDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("inward")]
    public string? Inward { get; init; }
    [JsonPropertyName("outward")]
    public string? Outward { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueLinkTypeDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueLinkTypeDtoJsonContext : JsonSerializerContext;