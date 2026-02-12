namespace TempoOutlookSync.Dto;

using System.Collections.Immutable;
using System.Text.Json.Serialization;

public sealed class TempoPlannerPayloadDto
{
    [JsonPropertyName("results")]
    public List<TempoPlannerEntryDto>? Results { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(TempoPlannerPayloadDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class TempoPlannerPayloadDtoJsonContext : JsonSerializerContext;