namespace TempoOutlookSync.Common;

using System.Text.Json.Serialization;

public sealed class Configuration
{
    public required string ApiToken { get; init; }
    public required string UserId { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(Configuration), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class ConfigurationJsonContext : JsonSerializerContext;