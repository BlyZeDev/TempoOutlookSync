namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraUserAvatarUrlsDto
{
    [JsonPropertyName("48x48")]
    public string? Avatar48 { get; init; }
    [JsonPropertyName("32x32")]
    public string? Avatar32 { get; init; }
    [JsonPropertyName("24x24")]
    public string? Avatar24 { get; init; }
    [JsonPropertyName("16x16")]
    public string? Avatar16 { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraUserAvatarUrlsDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraUserAvatarDtoJsonContext : JsonSerializerContext;