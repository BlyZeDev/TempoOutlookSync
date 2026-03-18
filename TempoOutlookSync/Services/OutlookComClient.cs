namespace TempoOutlookSync.Services;

using Microsoft.Office.Interop.Outlook;
using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using TempoOutlookSync.Models;

public sealed class OutlookComClient : IDisposable
{
    private const string OutlookTempoIdProperty = "TempoId";
    private const string OutlookTempoUpdatedProperty = "TempoUpdated";
    private const string OutlookJiraUpdatedProperty = "JiraUpdated";

    private readonly Thread _outlookThread;
    private readonly BlockingCollection<System.Action> _queue;

    public OutlookComClient()
    {
        _queue = [];

        _outlookThread = new Thread(() =>
        {
            foreach (var action in _queue.GetConsumingEnumerable())
            {
                action();
            }
        });
        _outlookThread.IsBackground = true;
        _outlookThread.SetApartmentState(ApartmentState.STA);
        _outlookThread.Start();
    }

    public HashSet<OutlookAppointmentRef> GetOutlookTempoAppointments()
    {
        return ExecuteSTA(() =>
        {
            var results = new HashSet<OutlookAppointmentRef>();

            Application? outlook = null;
            NameSpace? ns = null;
            MAPIFolder? folder = null;
            Items? items = null;
            Items? restricted = null;

            try
            {
                outlook = GetApplication();
                ns = outlook.GetNamespace("MAPI");
                folder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                items = folder.Items;

                items.IncludeRecurrences = false;
                restricted = items.Restrict($"@SQL=\"http://schemas.microsoft.com/mapi/string/{{00020329-0000-0000-C000-000000000046}}/{OutlookTempoIdProperty}\" IS NOT NULL");
                restricted.Sort("[Start]");

                for (int i = 1; i <= restricted.Count; i++)
                {
                    var item = restricted[i] as AppointmentItem;

                    if (item is not null)
                    {
                        var userProps = item.UserProperties;

                        var tempoIdProp = userProps.Find(OutlookTempoIdProperty);
                        var tempoUpdatedProp = userProps.Find(OutlookTempoUpdatedProperty);
                        var jiraUpdatedProp = userProps.Find(OutlookJiraUpdatedProperty);

                        if (tempoIdProp?.Value is string tempoIdStr)
                        {
                            if (int.TryParse(tempoIdStr, out var tempoId))
                            {
                                results.Add(new OutlookAppointmentRef
                                {
                                    TempoId = tempoId,
                                    EntryId = item.EntryID,
                                    Start = item.Start,
                                    End = item.End,
                                    TempoUpdated = ParseDateTime(tempoUpdatedProp?.Value),
                                    JiraUpdated = ParseDateTime(jiraUpdatedProp?.Value)
                                });
                            }
                        }

                        ReleaseComObject(jiraUpdatedProp);
                        ReleaseComObject(tempoUpdatedProp);
                        ReleaseComObject(tempoIdProp);
                        ReleaseComObject(userProps);
                    }

                    ReleaseComObject(item);
                }
            }
            finally
            {
                ReleaseComObject(restricted);
                ReleaseComObject(items);
                ReleaseComObject(folder);
                ReleaseComObject(ns);
                ReleaseComObject(outlook);
            }

            return results;
        });
    }

    public void DeleteByEntryId(string entryId)
    {
        ExecuteSTA(() =>
        {
            Application? outlook = null;
            NameSpace? ns = null;

            try
            {
                outlook = GetApplication();
                ns = outlook.GetNamespace("MAPI");

                var item = ns.GetItemFromID(entryId) as AppointmentItem;
                item?.Delete();
                ReleaseComObject(item);
            }
            finally
            {
                ReleaseComObject(ns);
                ReleaseComObject(outlook);
            }

            return 0;
        });
    }

    public void SaveNonRecurring(OutlookAppointmentInfo info)
    {
        ExecuteSTA(() =>
        {
            for (var day = info.TempoEntry.Start.Date; day <= info.TempoEntry.End.Date; day = day.AddDays(1))
            {
                if (!info.TempoEntry.IncludeNonWorkingDays && (day.DayOfWeek is DayOfWeek.Saturday || day.DayOfWeek is DayOfWeek.Sunday)) continue;

                CreateSingle(info, day + info.TempoEntry.StartTime);
            }

            return 0;
        });
    }

    public void SaveWeeklyRecurring(OutlookAppointmentInfo info)
    {
        ExecuteSTA(() =>
        {
            var baseStart = info.TempoEntry.Start.Date + info.TempoEntry.StartTime;
            var appointment = CreateBase(info, baseStart);

            ApplyRecurrence(appointment, info.TempoEntry, baseStart);

            appointment.Save();
            ReleaseComObject(appointment);

            return 0;
        });
    }

    public void SaveMonthlyRecurrence(OutlookAppointmentInfo info)
    {
        ExecuteSTA(() =>
        {
            for (var day = info.TempoEntry.Start.Date; day <= info.TempoEntry.End.Date; day = day.AddDays(1))
            {
                var monthlyStart = day + info.TempoEntry.StartTime;
                var monthlyAppointment = CreateBase(info, monthlyStart);

                ApplyRecurrence(monthlyAppointment, info.TempoEntry, monthlyStart);

                monthlyAppointment.Save();
                ReleaseComObject(monthlyAppointment);
            }

            return 0;
        });
    }

    public void Dispose()
    {
        _queue.CompleteAdding();
        _outlookThread.Join();
        _queue.Dispose();
    }

    private TResult ExecuteSTA<TResult>(Func<TResult> func)
    {
        var tcs = new TaskCompletionSource<TResult>();

        _queue.Add(() =>
        {
            try
            {
                var result = func();
                tcs.SetResult(result);
            }
            catch (System.Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        
        return tcs.Task.GetAwaiter().GetResult();
    }

    private static AppointmentItem CreateBase(OutlookAppointmentInfo info, DateTime start)
    {
        Application? outlook = null;

        try
        {
            outlook = GetApplication();

            var appointment = (AppointmentItem)outlook.CreateItem(OlItemType.olAppointmentItem);

            appointment.Subject = info.Subject;
            appointment.BodyFormat = OlBodyFormat.olFormatRichText;

            appointment.Body = BuildAppointmentRtf(info.Subject, info.Summary, info.Url);

            appointment.Start = start;
            appointment.BusyStatus = OlBusyStatus.olBusy;
            appointment.ReminderSet = false;

            if (info.Category is not null)
            {
                var ns = outlook.GetNamespace("MAPI");
                var categories = ns.Categories;

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

                ReleaseComObject(categories);
                ReleaseComObject(ns);
            }

            appointment.UserProperties.Add(OutlookTempoIdProperty, OlUserPropertyType.olText, true).Value = info.TempoEntry.Id.ToString();
            appointment.UserProperties.Add(OutlookTempoUpdatedProperty, OlUserPropertyType.olText, true).Value = info.TempoEntry.LastUpdated.ToString("O");

            if (info.LastUpdated.HasValue) appointment.UserProperties.Add(OutlookJiraUpdatedProperty, OlUserPropertyType.olText, true).Value = info.LastUpdated.Value.ToString("O");

            return appointment;
        }
        finally
        {
            ReleaseComObject(outlook);
        }
    }

    private static void CreateSingle(OutlookAppointmentInfo info, DateTime start)
    {
        var appointment = CreateBase(info, start);
        appointment.End = start + info.TempoEntry.DurationPerDay;

        appointment.Save();
        ReleaseComObject(appointment);
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

    private static DateTime? ParseDateTime(object? value)
    {
        if (value is not string str) return null;
        if (!DateTimeOffset.TryParse(str, out var date)) return null;
        return date.UtcDateTime;
    }

    private static Application GetApplication()
    {
        ExceptionDispatchInfo? info = null;

        for (int i = 0; i < 3; i++)
        {
            try
            {
                info = null;
                return new Application();
            }
            catch (COMException ex) when ((uint)ex.ErrorCode is 0x80010001 or 0x8001010A or 0x800706BA or 0x80010108)
            {
                info = ExceptionDispatchInfo.Capture(ex);
            }
        }

        throw info?.SourceException ?? new InvalidComObjectException("Failed to create Outlook.Application");
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
}