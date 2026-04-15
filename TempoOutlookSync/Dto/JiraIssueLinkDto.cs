namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraIssueLinkDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("type")]
    public required JiraIssueLinkTypeDto Type { get; init; }
    [JsonPropertyName("inwardIssue")]
    public JiraLinkedIssueDto? InwardIssue { get; init; }
    [JsonPropertyName("outwardIssue")]
    public JiraLinkedIssueDto? OutwardIssue { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraIssueLinkDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraIssueLinkDtoJsonContext : JsonSerializerContext;