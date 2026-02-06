using System;
using System.Text.Json;

namespace TempoOutlookSync
{
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

        public TempoPlannerEntry(int id, DateTime start, DateTime end, string description, TimeSpan startTime, TimeSpan durationPerDay, RecurrenceRule recurrenceRule, DateTime recurrenceEnd, bool includeWeekends, DateTime lastUpdated)
        {
            Id = id;
            Start = start;
            End = end;
            Description = description;
            StartTime = startTime;
            DurationPerDay = durationPerDay;
            RecurrenceRule = recurrenceRule;
            RecurrenceEnd = recurrenceEnd;
            IncludeNonWorkingDays = includeWeekends;
            LastUpdated = lastUpdated;
        }

        public override string ToString() => JsonSerializer.Serialize(this, _options);
    }
}
