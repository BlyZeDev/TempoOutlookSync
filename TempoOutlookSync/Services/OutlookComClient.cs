namespace TempoOutlookSync.Services;

using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using TempoOutlookSync.Models;

using OutlookCom = Microsoft.Office.Interop.Outlook;

public sealed class OutlookComClient : IDisposable
{
    private const string OutlookTempoIdProperty = "TempoId";
    private const string OutlookTempoUpdatedProperty = "TempoUpdated";
    private const string OutlookJiraUpdatedProperty = "JiraUpdated";

    private const string TempoSQLFilter = $"@SQL=\"http://schemas.microsoft.com/mapi/string/{{00020329-0000-0000-C000-000000000046}}/{OutlookTempoIdProperty}\" IS NOT NULL";

    private const uint RPC_E_CALL_REJECTED = 0x80010001;
    private const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;
    private const uint RPC_E_DISCONNECTED = 0x80010108;
    private const uint RPC_E_SERVER_UNAVAILABLE = 0x800706BA;

    private readonly TempoOutlookSyncContext _context;
    private readonly UpdateHandler _update;

    private readonly Thread _comThread;
    private readonly BlockingCollection<Action> _queue;

    public OutlookComClient(TempoOutlookSyncContext context, UpdateHandler update)
    {
        _context = context;
        _update = update;

        _queue = [];

        _comThread = new Thread(() =>
        {
            foreach (var action in _queue.GetConsumingEnumerable())
            {
                action();
            }
        });
        _comThread.IsBackground = true;
        _comThread.SetApartmentState(ApartmentState.STA);
        _comThread.Start();
    }

    public HashSet<OutlookAppointmentRef> GetTempoAppointments()
    {
        return ExecuteSTA(() =>
        {
            var results = new HashSet<OutlookAppointmentRef>();

            OutlookCom.Application? outlook = null;
            OutlookCom.NameSpace? ns = null;
            OutlookCom.MAPIFolder? folder = null;
            OutlookCom.Items? items = null;
            OutlookCom.Items? restricted = null;

            try
            {
                outlook = GetOutlookApp();
                ns = outlook.GetNamespace("MAPI");
                folder = ns.GetDefaultFolder(OutlookCom.OlDefaultFolders.olFolderCalendar);
                items = folder.Items;

                items.IncludeRecurrences = false;
                restricted = items.Restrict(TempoSQLFilter);
                restricted.Sort("[Start]");

                for (int i = 1; i <= restricted.Count; i++)
                {
                    var item = restricted[i] as OutlookCom.AppointmentItem;

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
            OutlookCom.Application? outlook = null;
            OutlookCom.NameSpace? ns = null;

            try
            {
                outlook = GetOutlookApp();
                ns = outlook.GetNamespace("MAPI");

                var item = ns.GetItemFromID(entryId) as OutlookCom.AppointmentItem;
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

    public void PurgeTrashedTempoAppointments()
    {
        ExecuteSTA(() =>
        {
            OutlookCom.Application? outlook = null;
            OutlookCom.NameSpace? ns = null;
            OutlookCom.MAPIFolder? folder = null;
            OutlookCom.Items? items = null;
            OutlookCom.Items? restricted = null;

            try
            {
                outlook = GetOutlookApp();
                ns = outlook.GetNamespace("MAPI");
                folder = ns.GetDefaultFolder(OutlookCom.OlDefaultFolders.olFolderDeletedItems);

                var userProps = folder.UserDefinedProperties;
                var existingProp = userProps.Find(OutlookTempoIdProperty);
                if (existingProp is null) userProps.Add(OutlookTempoIdProperty, OutlookCom.OlUserPropertyType.olText);

                ReleaseComObject(userProps);
                ReleaseComObject(existingProp);

                items = folder.Items;
                restricted = items.Restrict(TempoSQLFilter);

                for (int i = restricted.Count; i >= 1; i--)
                {
                    var item = restricted[i] as OutlookCom.AppointmentItem;
                    item?.Delete();
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

            return 0;
        });
    }

    public void SaveNonRecurring(OutlookAppointmentCreationInfo info)
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

    public void SaveWeeklyRecurring(OutlookAppointmentCreationInfo info)
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

    public void SaveMonthlyRecurrence(OutlookAppointmentCreationInfo info)
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
        _comThread.Join();
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
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });

        return tcs.Task.GetAwaiter().GetResult();
    }

    private OutlookCom.AppointmentItem CreateBase(OutlookAppointmentCreationInfo info, DateTime start)
    {
        OutlookCom.Application? outlook = null;

        try
        {
            outlook = GetOutlookApp();

            var appointment = (OutlookCom.AppointmentItem)outlook.CreateItem(OutlookCom.OlItemType.olAppointmentItem);

            appointment.Subject = info.Subject;
            appointment.BodyFormat = OutlookCom.OlBodyFormat.olFormatHTML;

            OutlookCom.NameSpace? ns = null;
            OutlookCom.MAPIFolder? deletedFolder = null;
            OutlookCom.Items? deletedFolderItems = null;
            OutlookCom.MailItem? mail = null;
            OutlookCom.Inspector? mailInspector = null;
            OutlookCom.Inspector? appointmentInspector = null;
            try
            {
                ns = outlook.GetNamespace("MAPI");

                deletedFolder = ns.GetDefaultFolder(OutlookCom.OlDefaultFolders.olFolderDeletedItems);
                deletedFolderItems = deletedFolder.Items;
                mail = (OutlookCom.MailItem)deletedFolderItems.Add(OutlookCom.OlItemType.olMailItem);
                mail.HTMLBody = info.BuildHtmlBody(_update.Version);

                mailInspector = mail.GetInspector;
                appointmentInspector = appointment.GetInspector;

                var mailDoc = mailInspector.WordEditor;
                var appointmentDoc = appointmentInspector.WordEditor;

                appointmentDoc.Range().FormattedText = mailDoc.Range().FormattedText;

                appointmentDoc.ShowSpellingErrors = false;
                appointmentDoc.Saved = true;

                mail.Delete();
            }
            finally
            {
                ReleaseComObject(appointmentInspector);
                ReleaseComObject(mailInspector);
                ReleaseComObject(mail);
                ReleaseComObject(deletedFolderItems);
                ReleaseComObject(deletedFolder);
            }

            appointment.Start = start;
            appointment.BusyStatus = OutlookCom.OlBusyStatus.olBusy;
            appointment.ReminderSet = false;

            if (info.Category is not null)
            {
                ns = outlook.GetNamespace("MAPI");
                var categories = ns.Categories;

                var category = categories[info.Category.Name];

                if (category is null) categories.Add(info.Category.Name, info.Category.Color, OutlookCom.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                else if (category.Color != info.Category.Color) category.Color = info.Category.Color;

                appointment.Categories = info.Category.Name;

                ReleaseComObject(category);
                ReleaseComObject(categories);
                ReleaseComObject(ns);
            }

            appointment.UserProperties.Add(OutlookTempoIdProperty, OutlookCom.OlUserPropertyType.olText, true).Value = info.TempoEntry.Id.ToString();
            appointment.UserProperties.Add(OutlookTempoUpdatedProperty, OutlookCom.OlUserPropertyType.olText, true).Value = info.TempoEntry.LastUpdated.ToString("O");

            if (info.LastUpdated.HasValue) appointment.UserProperties.Add(OutlookJiraUpdatedProperty, OutlookCom.OlUserPropertyType.olText, true).Value = info.LastUpdated.Value.ToString("O");

            return appointment;
        }
        finally
        {
            ReleaseComObject(outlook);
        }
    }

    private void CreateSingle(OutlookAppointmentCreationInfo info, DateTime start)
    {
        var appointment = CreateBase(info, start);
        appointment.End = start + info.TempoEntry.DurationPerDay;

        appointment.Save();
        ReleaseComObject(appointment);
    }

    private static void ApplyRecurrence(OutlookCom.AppointmentItem appointment, TempoPlannerEntry entry, DateTime start)
    {
        var recurrence = appointment.GetRecurrencePattern();

        if (entry.RecurrenceRule is TempoRecurrenceRule.Monthly)
        {
            recurrence.RecurrenceType = OutlookCom.OlRecurrenceType.olRecursMonthly;
            recurrence.DayOfMonth = entry.Start.Day;
        }
        else
        {
            recurrence.RecurrenceType = OutlookCom.OlRecurrenceType.olRecursWeekly;
            recurrence.Interval = entry.RecurrenceRule is TempoRecurrenceRule.BiWeekly ? 2 : 1;
            recurrence.DayOfWeekMask = BuildMask(entry.Start.Date, entry.End.Date, entry.IncludeNonWorkingDays);
        }

        recurrence.NoEndDate = false;

        recurrence.PatternStartDate = entry.Start.Date;
        recurrence.PatternEndDate = entry.RecurrenceEnd.Date;

        recurrence.StartTime = start;
        recurrence.EndTime = start + entry.DurationPerDay;
    }

    private static OutlookCom.OlDaysOfWeek BuildMask(DateTime start, DateTime end, bool includeNonWorkingDays)
    {
        OutlookCom.OlDaysOfWeek mask = 0;

        for (var date = start; date <= end; date = date.AddDays(1))
        {
            switch (date.DayOfWeek)
            {
                case DayOfWeek.Monday: mask |= OutlookCom.OlDaysOfWeek.olMonday; break;
                case DayOfWeek.Tuesday: mask |= OutlookCom.OlDaysOfWeek.olTuesday; break;
                case DayOfWeek.Wednesday: mask |= OutlookCom.OlDaysOfWeek.olWednesday; break;
                case DayOfWeek.Thursday: mask |= OutlookCom.OlDaysOfWeek.olThursday; break;
                case DayOfWeek.Friday: mask |= OutlookCom.OlDaysOfWeek.olFriday; break;
                case DayOfWeek.Saturday: mask |= OutlookCom.OlDaysOfWeek.olSaturday; break;
                case DayOfWeek.Sunday: mask |= OutlookCom.OlDaysOfWeek.olSunday; break;
            }
        }

        if (!includeNonWorkingDays)
        {
            mask &= ~OutlookCom.OlDaysOfWeek.olSaturday;
            mask &= ~OutlookCom.OlDaysOfWeek.olSunday;
        }

        return mask;
    }

    private static DateTime? ParseDateTime(object? value)
    {
        if (value is not string str) return null;
        if (!DateTimeOffset.TryParse(str, out var date)) return null;
        return date.UtcDateTime;
    }

    private static OutlookCom.Application GetOutlookApp()
    {
        ExceptionDispatchInfo? info = null;

        for (int i = 0; i < 3; i++)
        {
            try
            {
                info = null;
                return new OutlookCom.Application();
            }
            catch (COMException ex) when ((uint)ex.ErrorCode is RPC_E_CALL_REJECTED or RPC_E_SERVERCALL_RETRYLATER or RPC_E_SERVER_UNAVAILABLE or RPC_E_DISCONNECTED)
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