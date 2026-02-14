namespace TempoOutlookSync.Dto;

using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

public sealed record TempoPlannerEntryDto
{
    [JsonPropertyName("id")]
    public required int Id { get; init; }
    [JsonPropertyName("startDate")]
    public required string StartDate { get; init; }
    [JsonPropertyName("endDate")]
    public required string EndDate { get; init; }
    [JsonPropertyName("description")]
    public string? Description { get; init; }
    [JsonPropertyName("startTime")]
    public required string StartTime { get; init; }
    [JsonPropertyName("plannedSecondsPerDay")]
    public required long PlannedSecondsPerDay { get; init; }
    [JsonPropertyName("rule")]
    public string? Rule { get; init; }
    [JsonPropertyName("recurrenceEndDate")]
    public required string RecurrenceEndDate { get; init; }
    [JsonPropertyName("includeNonWorkingDays")]
    public bool? IncludeNonWorkingDays { get; init; }
    [JsonPropertyName("createdAt")]
    public required string CreatedAt { get; init; }
    [JsonPropertyName("updatedAt")]
    public string? UpdatedAt { get; init; }
    [JsonPropertyName("planItem")]
    public TempoPlannerEntryPlanItemDto? PlanItem { get; init; }
}

[JsonSourceGenerationOptions(WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(TempoPlannerEntryDto), GenerationMode = JsonSourceGenerationMode.Default)]
public sealed partial class TempoPlannerEntryDtoJsonContext : JsonSerializerContext;