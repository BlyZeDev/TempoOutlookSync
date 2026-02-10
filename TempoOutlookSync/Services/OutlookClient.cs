namespace TempoOutlookSync.Services;

using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using TempoOutlookSync.Common;

public sealed class OutlookClient : IDisposable
{
    private const string OutlookTempoIdProperty = "TempoId";
    private const string OutlookTempoUpdatedProperty = "TempoUpdated";

    private readonly Application _outlook;

    public OutlookClient() => _outlook = new Application();

    public Dictionary<string, HashSet<AppointmentItem>> GetOutlookTempoAppointments(DateTime start)
    {
        var items = _outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items;
        items.IncludeRecurrences = false;
        items = items.Restrict($"@SQL=\"http://schemas.microsoft.com/mapi/string/{{00020329-0000-0000-C000-000000000046}}/{OutlookTempoIdProperty}\" IS NOT NULL");
        items.Sort("[Start]");

        return items.OfType<AppointmentItem>()
            .Select(item => new
            {
                Item = item,
                TempoId = item.UserProperties.Find(OutlookTempoIdProperty)
            })
            .Where(x => x.TempoId?.Value is not null)
            .GroupBy(x => (string)Convert.ToString(x.TempoId.Value))
            .ToDictionary(key => key.Key, value => new HashSet<AppointmentItem>(value.Select(x => x.Item)));
    }

    public bool DeleteIfOutdated(AppointmentItem appointment, DateTime latestUpdate)
    {
        var property = appointment.UserProperties.Find(OutlookTempoUpdatedProperty)?.Value;

        if (property == null)
        {
            appointment.Delete();
            return true;
        }

        var lastUpdated = ((DateTimeOffset)DateTimeOffset.Parse(property)).UtcDateTime;
        if (lastUpdated < latestUpdate)
        {
            appointment.Delete();
            return true;
        }
        else return false;
    }

    public void SaveNonRecurring(TempoPlannerEntry entry)
    {
        for (var day = entry.Start.Date; day <= entry.End.Date; day = day.AddDays(1))
        {
            if (!entry.IncludeNonWorkingDays && (day.DayOfWeek is DayOfWeek.Saturday || day.DayOfWeek is DayOfWeek.Sunday)) continue;

            CreateSingle(_outlook, entry, day + entry.StartTime);
        }
    }

    public void SaveWeeklyRecurring(TempoPlannerEntry entry)
    {
        var baseStart = entry.Start.Date + entry.StartTime;
        var appointment = CreateBase(_outlook, entry, baseStart);
        ApplyRecurrence(appointment, entry, baseStart);
        appointment.Save();
    }

    public void SaveMonthlyRecurrence(TempoPlannerEntry entry)
    {
        for (var day = entry.Start.Date; day <= entry.End.Date; day = day.AddDays(1))
        {
            var monthlyStart = day + entry.StartTime;

            var monthlyAppointment = CreateBase(_outlook, entry, monthlyStart);

            ApplyRecurrence(monthlyAppointment, entry, monthlyStart);

            monthlyAppointment.Save();
        }
    }

    public void Dispose()
    {
        Marshal.ReleaseComObject(_outlook);
    }

    private static AppointmentItem CreateBase(Application outlook, TempoPlannerEntry entry, DateTime start)
    {
        var appointment = (AppointmentItem)outlook.CreateItem(OlItemType.olAppointmentItem);

        appointment.Subject = entry.Description;
        appointment.Body = $"[AutoImport by Jira Tempo]\n{entry.Description}";
        appointment.Start = start;
        appointment.BusyStatus = OlBusyStatus.olBusy;
        appointment.ReminderSet = false;

        appointment.UserProperties.Add(OutlookTempoIdProperty, OlUserPropertyType.olText, true).Value = entry.Id.ToString();
        appointment.UserProperties.Add(OutlookTempoUpdatedProperty, OlUserPropertyType.olText, true).Value = entry.LastUpdated.ToString("O");

        return appointment;
    }

    private static void CreateSingle(Application outlook, TempoPlannerEntry entry, DateTime start)
    {
        var appointment = CreateBase(outlook, entry, start);
        appointment.End = start + entry.DurationPerDay;
        appointment.Save();
    }

    private static void ApplyRecurrence(AppointmentItem appointment, TempoPlannerEntry entry, DateTime start)
    {
        var recurrence = appointment.GetRecurrencePattern();

        if (entry.RecurrenceRule is RecurrenceRule.Monthly)
        {
            recurrence.RecurrenceType = OlRecurrenceType.olRecursMonthly;
            recurrence.DayOfMonth = entry.Start.Day;
        }
        else
        {
            recurrence.RecurrenceType = OlRecurrenceType.olRecursWeekly;
            recurrence.Interval = entry.RecurrenceRule is RecurrenceRule.BiWeekly ? 2 : 1;
            recurrence.DayOfWeekMask = BuildMask(entry.Start.Date, entry.End.Date, entry.IncludeNonWorkingDays);
        }

        recurrence.NoEndDate = false;

        recurrence.PatternStartDate = entry.Start.Date;
        recurrence.PatternEndDate = entry.RecurrenceEnd.Date;

        recurrence.StartTime = start;
        recurrence.EndTime = start + entry.DurationPerDay;
    }

    private static OlDaysOfWeek BuildMask(DateTime start, DateTime end, bool includeNonWorkingDays)
    {
        OlDaysOfWeek mask = 0;

        for (var date = start; date <= end; date = date.AddDays(1))
        {
            switch (date.DayOfWeek)
            {
                case DayOfWeek.Monday: mask |= OlDaysOfWeek.olMonday; break;
                case DayOfWeek.Tuesday: mask |= OlDaysOfWeek.olTuesday; break;
                case DayOfWeek.Wednesday: mask |= OlDaysOfWeek.olWednesday; break;
                case DayOfWeek.Thursday: mask |= OlDaysOfWeek.olThursday; break;
                case DayOfWeek.Friday: mask |= OlDaysOfWeek.olFriday; break;
                case DayOfWeek.Saturday: mask |= OlDaysOfWeek.olSaturday; break;
                case DayOfWeek.Sunday: mask |= OlDaysOfWeek.olSunday; break;
            }
        }

        if (!includeNonWorkingDays)
        {
            mask &= ~OlDaysOfWeek.olSaturday;
            mask &= ~OlDaysOfWeek.olSunday;
        }

        return mask;
    }
}
