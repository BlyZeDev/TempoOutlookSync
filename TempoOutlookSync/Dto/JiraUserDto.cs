namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record JiraUserDto
{
    [JsonPropertyName("emailAddress")]
    public string? EmailAddress { get; init; }
    [JsonPropertyName("displayName")]
    public string? DisplayName { get; init; }
    [JsonPropertyName("avatarUrls")]
    public JiraUserAvatarUrlsDto? AvatarUrls { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(JiraUserDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class JiraUserDtoJsonContext : JsonSerializerContext;