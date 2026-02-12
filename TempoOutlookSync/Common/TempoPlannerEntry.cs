namespace TempoOutlookSync.Common;

using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using TempoOutlookSync.Dto;
using TempoOutlookSync.Services;

public sealed class TempoPlannerEntry
{
    private static readonly JsonSerializerOptions _options = new JsonSerializerOptions()
    {
        WriteIndented = true
    };

    public int Id { get; }
    public DateTime Start { get; }
    public DateTime End { get; }
    public string Description { get; }
    public TimeSpan StartTime { get; }
    public TimeSpan DurationPerDay { get; }
    public RecurrenceRule RecurrenceRule { get; }
    public DateTime RecurrenceEnd { get; }
    public bool IncludeNonWorkingDays { get; }
    public DateTime LastUpdated { get; }

    [JsonConstructor]
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

        Description = dto.Description ?? $"Issue #{Id}";

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
    }

    public override string ToString() => JsonSerializer.Serialize(this, _options);

    private static RecurrenceRule ParseRecurrenceRule(string recurrenceRule)
    {
        switch (recurrenceRule.ToLower())
        {
            case "weekly": return RecurrenceRule.Weekly;
            case "bi_weekly": return RecurrenceRule.BiWeekly;
            case "monthly": return RecurrenceRule.Monthly;
        }

        return RecurrenceRule.Never;
    }
}
