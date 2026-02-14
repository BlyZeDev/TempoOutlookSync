namespace TempoOutlookSync.Models;

using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using TempoOutlookSync.Dto;
using TempoOutlookSync.Services;

public sealed record TempoPlannerEntry
{
    private static readonly JsonSerializerOptions _options = new JsonSerializerOptions()
    {
        WriteIndented = true
    };

    public int Id { get; }
    public DateTime Start { get; }
    public DateTime End { get; }
    public string? Description { get; }
    public TimeSpan StartTime { get; }
    public TimeSpan DurationPerDay { get; }
    public TempoRecurrenceRule RecurrenceRule { get; }
    public DateTime RecurrenceEnd { get; }
    public bool IncludeNonWorkingDays { get; }
    public DateTime LastUpdated { get; }
    public string? PlanItemId { get; }
    public TempoPlanItemType PlanItemType { get; }

    public TempoPlannerEntry(TempoPlannerEntryDto dto)
    {
        Id = dto.Id;

        Start = DateTime.ParseExact(
            dto.StartDate ?? throw new JsonException("startDate missing"),
            TempoApiClient.TempoDateFormat,
            CultureInfo.InvariantCulture);

        End = DateTime.ParseExact(
            dto.EndDate ?? throw new JsonException("endDate missing"),
            TempoApiClient.TempoDateFormat,
            CultureInfo.InvariantCulture);

        Description = dto.Description;

        StartTime = TimeSpan.ParseExact(
            dto.StartTime ?? throw new JsonException("startTime missing"),
            @"hh\:mm",
            CultureInfo.InvariantCulture);

        DurationPerDay = TimeSpan.FromSeconds(dto.PlannedSecondsPerDay);

        RecurrenceRule = ParseRecurrenceRule(
            dto.Rule ?? throw new JsonException("rule missing"));

        RecurrenceEnd = DateTime.ParseExact(
            dto.RecurrenceEndDate ?? throw new JsonException("recurrenceEndDate missing"),
            TempoApiClient.TempoDateFormat,
            CultureInfo.InvariantCulture);

        IncludeNonWorkingDays = dto.IncludeNonWorkingDays ?? true;

        LastUpdated = DateTimeOffset.Parse(dto.UpdatedAt ?? dto.CreatedAt).UtcDateTime;

        PlanItemId = dto.PlanItem?.Id;

        PlanItemType = ParsePlanItemType(dto.PlanItem?.Type);
    }

    public override string ToString() => JsonSerializer.Serialize(this, _options);

    private static TempoRecurrenceRule ParseRecurrenceRule(string recurrenceRule)
    {
        return recurrenceRule switch
        {
            var rule when rule.Equals("weekly", StringComparison.OrdinalIgnoreCase) => TempoRecurrenceRule.Weekly,
            var rule when rule.Equals("bi_weekly", StringComparison.OrdinalIgnoreCase) => TempoRecurrenceRule.BiWeekly,
            var rule when rule.Equals("monthly", StringComparison.OrdinalIgnoreCase) => TempoRecurrenceRule.Monthly,
            _ => TempoRecurrenceRule.Never
        };
    }

    private static TempoPlanItemType ParsePlanItemType(string? planItemType)
    {
        return planItemType switch
        {
            null => TempoPlanItemType.Unknown,
            var type when type.Equals("issue", StringComparison.OrdinalIgnoreCase) => TempoPlanItemType.Issue,
            var type when type.Equals("project", StringComparison.OrdinalIgnoreCase) => TempoPlanItemType.Project,
            _ => TempoPlanItemType.Unknown
        };
    }
}
