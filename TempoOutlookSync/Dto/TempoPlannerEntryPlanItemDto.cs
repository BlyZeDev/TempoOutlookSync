namespace TempoOutlookSync.Dto;

using System.Text.Json.Serialization;

public sealed record TempoPlannerEntryPlanItemDto
{

}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(TempoPlannerEntryPlanItemDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class TempoPlannerEntryPlanItemDtoJsonContext : JsonSerializerContext;