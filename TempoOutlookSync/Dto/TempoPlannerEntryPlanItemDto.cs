namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record TempoPlannerEntryPlanItemDto
{
    [JsonPropertyName("id")]
    public required string Id { get; init; }
    [JsonPropertyName("type")]
    public string? Type { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(TempoPlannerEntryPlanItemDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class TempoPlannerEntryPlanItemDtoJsonContext : JsonSerializerContext;