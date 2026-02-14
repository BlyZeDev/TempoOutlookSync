namespace TempoOutlookSync.Services;

using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Text;
using TempoOutlookSync.Models;

public sealed class OutlookClient : IDisposable
{
    private const string OutlookTempoIdProperty = "TempoId";
    private const string OutlookTempoUpdatedProperty = "TempoUpdated";
    private const string OutlookJiraUpdatedProperty = "JiraUpdated";

    private readonly Application _outlook;

    public OutlookClient() => _outlook = new Application();

    public Dictionary<int, HashSet<AppointmentItem>> GetOutlookTempoAppointments(DateTime start)
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
            .GroupBy(x => (int)int.Parse(x.TempoId.Value))
            .ToDictionary(key => key.Key, value => new HashSet<AppointmentItem>(value.Select(x => x.Item)));
    }

    public bool DeleteIfDifferent(AppointmentItem appointment, DateTime latestTempoUpdate, DateTime latestJiraUpdate)
    {
        var tempoUpdated = appointment.UserProperties.Find(OutlookTempoUpdatedProperty)?.Value;
        var jiraUpdated = appointment.UserProperties.Find(OutlookJiraUpdatedProperty)?.Value;

        if (tempoUpdated is null || jiraUpdated is null)
        {
            appointment.Delete();
            return true;
        }

        var lastTempoUpdate = ((DateTimeOffset)DateTimeOffset.Parse(tempoUpdated)).UtcDateTime;
        var lastJiraUpdate = ((DateTimeOffset)DateTimeOffset.Parse(jiraUpdated)).UtcDateTime;

        if (lastTempoUpdate == latestTempoUpdate && lastJiraUpdate == latestJiraUpdate) return false;

        appointment.Delete();
        return true;
    }

    public void SaveNonRecurring(OutlookAppointmentInfo info)
    {
        for (var day = info.TempoEntry.Start.Date; day <= info.TempoEntry.End.Date; day = day.AddDays(1))
        {
            if (!info.TempoEntry.IncludeNonWorkingDays && (day.DayOfWeek is DayOfWeek.Saturday || day.DayOfWeek is DayOfWeek.Sunday)) continue;

            CreateSingle(_outlook, info, day + info.TempoEntry.StartTime);
        }
    }

    public void SaveWeeklyRecurring(OutlookAppointmentInfo info)
    {
        var baseStart = info.TempoEntry.Start.Date + info.TempoEntry.StartTime;
        var appointment = CreateBase(_outlook, info, baseStart);
        ApplyRecurrence(appointment, info.TempoEntry, baseStart);
        appointment.Save();
    }

    public void SaveMonthlyRecurrence(OutlookAppointmentInfo info)
    {
        for (var day = info.TempoEntry.Start.Date; day <= info.TempoEntry.End.Date; day = day.AddDays(1))
        {
            var monthlyStart = day + info.TempoEntry.StartTime;

            var monthlyAppointment = CreateBase(_outlook, info, monthlyStart);

            ApplyRecurrence(monthlyAppointment, info.TempoEntry, monthlyStart);

            monthlyAppointment.Save();
        }
    }

    public void Dispose()
    {
        Marshal.ReleaseComObject(_outlook);
    }

    private static AppointmentItem CreateBase(Application outlook, OutlookAppointmentInfo info, DateTime start)
    {
        var appointment = (AppointmentItem)outlook.CreateItem(OlItemType.olAppointmentItem);

        appointment.Subject = info.Subject;
        appointment.BodyFormat = OlBodyFormat.olFormatRichText;

        appointment.Body = BuildAppointmentRtf(info.Subject, info.Summary, info.Url);

        appointment.Start = start;
        appointment.BusyStatus = OlBusyStatus.olBusy;
        appointment.ReminderSet = false;

        if (info.Category is not null)
        {
            var categories = outlook.GetNamespace("MAPI").Categories;
            Category? category;
            try
            {
                category = categories[info.Category.Name];
            }
            catch (System.Exception)
            {
                category = null;
            }

            if (category is null) categories.Add(info.Category.Name, info.Category.Color, OlCategoryShortcutKey.olCategoryShortcutKeyNone);

            appointment.Categories = info.Category.Name;
        }

        appointment.UserProperties.Add(OutlookTempoIdProperty, OlUserPropertyType.olText, true).Value = info.TempoEntry.Id.ToString();
        appointment.UserProperties.Add(OutlookTempoUpdatedProperty, OlUserPropertyType.olText, true).Value = info.TempoEntry.LastUpdated.ToString("O");

        if (info.LastUpdated.HasValue) appointment.UserProperties.Add(OutlookJiraUpdatedProperty, OlUserPropertyType.olText, true).Value = info.LastUpdated.Value.ToString("O");

        return appointment;
    }

    private static void CreateSingle(Application outlook, OutlookAppointmentInfo info, DateTime start)
    {
        var appointment = CreateBase(outlook, info, start);
        appointment.End = start + info.TempoEntry.DurationPerDay;
        appointment.Save();
    }

    private static void ApplyRecurrence(AppointmentItem appointment, TempoPlannerEntry entry, DateTime start)
    {
        var recurrence = appointment.GetRecurrencePattern();

        if (entry.RecurrenceRule is TempoRecurrenceRule.Monthly)
        {
            recurrence.RecurrenceType = OlRecurrenceType.olRecursMonthly;
            recurrence.DayOfMonth = entry.Start.Day;
        }
        else
        {
            recurrence.RecurrenceType = OlRecurrenceType.olRecursWeekly;
            recurrence.Interval = entry.RecurrenceRule is TempoRecurrenceRule.BiWeekly ? 2 : 1;
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

    private static string BuildAppointmentRtf(string subject, string summary, string? permalink)
    {
        var sb = new StringBuilder();

        sb.AppendLine(@"{\rtf1\ansi\deff0");
        sb.AppendLine(@"{\fonttbl{\f0\fnil\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}}");
        sb.AppendLine(@"{\colortbl ;\red46\green134\blue193;\red100\green100\blue100;}");
        sb.AppendLine(@"\viewkind4\uc1\pard\sl240\slmult1");

        sb.AppendLine(@"\f0\fs20\cf2\i Auto-imported from Jira Tempo\i0\cf0\par\par");

        if (!string.IsNullOrWhiteSpace(subject)) sb.AppendLine($@"\f0\fs26\cf0\b {EscapeRtf(subject)}\b0\par");

        sb.AppendLine(@"\fs22\par");

        if (!string.IsNullOrWhiteSpace(summary)) sb.AppendLine($@"{EscapeRtf(summary)}\par\par");

        if (!string.IsNullOrWhiteSpace(permalink))
        {
            var url = EscapeRtf(permalink);

            sb.AppendLine(@"\cf1");
            sb.AppendLine($@"{{\field{{\*\fldinst HYPERLINK ""{url}""}}{{\fldrslt\ul {url}\ulnone}}}}");
            sb.AppendLine(@"\cf0\par");
        }

        sb.AppendLine(@"\par\fs18\cf2 Please do not modify this appointment manually if it is synced automatically.\cf0");
        sb.AppendLine(@"}");

        return sb.ToString();
    }

    private static string EscapeRtf(string? value)
    {
        if (string.IsNullOrEmpty(value)) return "";

        return value.Replace(@"\", @"\\").Replace("{", @"\{").Replace("}", @"\}").Replace("\r\n", @"\par ").Replace("\n", @"\par ");
    }
}