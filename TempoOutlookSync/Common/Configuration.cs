namespace TempoOutlookSync.Common;

using System.Text.Json.Serialization;

public sealed class Configuration
{
    public required string Email { get; init; }
    public required string JiraApiToken { get; init; }
    public required string UserId { get; init; }
    public required string TempoApiToken { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(Configuration), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class ConfigurationJsonContext : JsonSerializerContext;