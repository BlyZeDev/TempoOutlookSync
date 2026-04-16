namespace TempoOutlookSync.Services;

using Microsoft.Office.Interop.Outlook;
using System.Collections.Concurrent;
using System.Net;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using TempoOutlookSync.Models;

public sealed class OutlookComClient : IDisposable
{
    private const string OutlookTempoIdProperty = "TempoId";
    private const string OutlookTempoUpdatedProperty = "TempoUpdated";
    private const string OutlookJiraUpdatedProperty = "JiraUpdated";

    private const string TempoSQLFilter = $"@SQL=\"http://schemas.microsoft.com/mapi/string/{{00020329-0000-0000-C000-000000000046}}/{OutlookTempoIdProperty}\" IS NOT NULL";

    private readonly UpdateHandler _update;

    private readonly Thread _outlookThread;
    private readonly BlockingCollection<System.Action> _queue;

    public OutlookComClient(UpdateHandler update)
    {
        _update = update;

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

    public HashSet<OutlookAppointmentRef> GetTempoAppointments()
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
                restricted = items.Restrict(TempoSQLFilter);
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

    public void PurgeTrashedTempoAppointments()
    {
        ExecuteSTA(() =>
        {
            Application? outlook = null;
            NameSpace? ns = null;
            MAPIFolder? folder = null;
            Items? items = null;
            Items? restricted = null;

            try
            {
                outlook = GetApplication();
                ns = outlook.GetNamespace("MAPI");
                folder = ns.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);

                var userProps = folder.UserDefinedProperties;
                var existingProp = userProps.Find(OutlookTempoIdProperty);
                if (existingProp is null) userProps.Add(OutlookTempoIdProperty, OlUserPropertyType.olText);

                ReleaseComObject(userProps);
                ReleaseComObject(existingProp);

                items = folder.Items;
                restricted = items.Restrict(TempoSQLFilter);

                for (int i = restricted.Count; i >= 1; i--)
                {
                    var item = restricted[i] as AppointmentItem;
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

    private AppointmentItem CreateBase(OutlookAppointmentCreationInfo info, DateTime start)
    {
        Application? outlook = null;

        try
        {
            outlook = GetApplication();

            var appointment = (AppointmentItem)outlook.CreateItem(OlItemType.olAppointmentItem);

            appointment.Subject = info.Subject;
            appointment.BodyFormat = OlBodyFormat.olFormatHTML;

            NameSpace? ns = null;
            MAPIFolder? deletedFolder = null;
            Items? deletedFolderItems = null;
            MailItem? mail = null;
            Inspector? mailInspector = null;
            Inspector? appointmentInspector = null;
            try
            {
                ns = outlook.GetNamespace("MAPI");

                deletedFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
                deletedFolderItems = deletedFolder.Items;
                mail = (MailItem)deletedFolderItems.Add(OlItemType.olMailItem);
                mail.HTMLBody = BuildAppointmentHtml(info);

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
            appointment.BusyStatus = OlBusyStatus.olBusy;
            appointment.ReminderSet = false;

            if (info.Category is not null)
            {
                ns = outlook.GetNamespace("MAPI");
                var categories = ns.Categories;

                var category = categories[info.Category.Name];

                if (category is null) categories.Add(info.Category.Name, info.Category.Color, OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                else if (category.Color != info.Category.Color) category.Color = info.Category.Color;

                appointment.Categories = info.Category.Name;

                ReleaseComObject(category);
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

    private void CreateSingle(OutlookAppointmentCreationInfo info, DateTime start)
    {
        var appointment = CreateBase(info, start);
        appointment.End = start + info.TempoEntry.DurationPerDay;

        appointment.Save();
        ReleaseComObject(appointment);
    }

    private string BuildAppointmentHtml(OutlookAppointmentCreationInfo info)
    {
        var sb = new StringBuilder();

        sb.AppendLine("""
            <div style="
                font-family:'Segoe UI', Calibri, sans-serif;
                color:#111111;
                font-size:14pt;
                line-height:20pt;
                mso-line-height-rule:exactly;
            ">
            """);

        sb.AppendLine("""
            <div style="color:#666666; font-size:11pt; font-style:italic;">
                Auto-imported from Jira Tempo
            </div>
            """);

        if (!string.IsNullOrWhiteSpace(info.Summary))
        {
            SetSpace(20);

            sb.AppendLine($"""
                <div style="font-size:20pt; font-weight:600;">
                    {WebUtility.HtmlEncode(info.Summary)}
                </div>
                """);
        }

        if (!string.IsNullOrWhiteSpace(info.Subject))
        {
            SetSpace(20);

            sb.AppendLine($"""
                <div style="font-size:15pt;">
                    {WebUtility.HtmlEncode(info.Subject)}
                </div>
                """);
        }

        if (!string.IsNullOrWhiteSpace(info.Url))
        {
            SetSpace(10);

            var url = WebUtility.HtmlEncode(info.Url);

            sb.AppendLine($"""
                <div>
                    <a href="{url}" style="color:#9B59B6; font-size:13pt; text-decoration:underline;">
                        {url}
                    </a>
                </div>
                """);
        }

        if (info.LinkedIssues.Count > 0)
        {
            SetSpace(30);

            sb.AppendLine($"""
                <div>
                    <div style="font-size:15pt; color:#2C3E50; font-weight:600;">
                        📎 Linked issues ({info.LinkedIssues.Count})
                    </div>
                """);

            foreach (var issue in info.LinkedIssues)
            {
                var relation = WebUtility.HtmlEncode(issue.RelationToBaseIssue);
                var summary = WebUtility.HtmlEncode(issue.LinkedIssue.Summary);
                var url = WebUtility.HtmlEncode(issue.LinkedIssue.Permalink);

                sb.AppendLine("<div>");

                if (!string.IsNullOrWhiteSpace(relation))
                {
                    sb.AppendLine($"""
                        <div style="font-size:11pt; color:#7F8C8D; font-style:italic;">
                            {relation}
                        </div>
                        """);
                }

                if (!string.IsNullOrWhiteSpace(summary))
                {
                    sb.AppendLine($"""
                        <div style="font-size:12pt;">
                            {summary}
                        </div>
                        """);
                }

                sb.AppendLine($"""
                        <div>
                            <a href="{url}" style="color:#9B59B6; font-size:12pt; text-decoration:underline;">
                                {url}
                            </a>
                        </div>
                    </div>
                    """);
            }

            sb.AppendLine("</div>");
        }

        if (!string.IsNullOrWhiteSpace(info.PlannedBy))
        {
            SetSpace(30);

            var avatarCell = "";

            if (!string.IsNullOrWhiteSpace(info.PlannedByAvatarUrl))
            {
                avatarCell = $"""
                    <td style="vertical-align:middle;">
                        <img src="{WebUtility.HtmlEncode(info.PlannedByAvatarUrl)}"
                             width="32" height="32"
                             style="display:block;" />
                    </td>
                    <td style="width:6px;"></td>
                    """;
            }

            sb.AppendLine($"""
                <table role="presentation" cellpadding="0" cellspacing="0">
                    <tr>
                        {avatarCell}
                        <td style="vertical-align:middle;">
                            Planned by {WebUtility.HtmlEncode(info.PlannedBy)}
                        </td>
                    </tr>
                </table>
                """);
        }

        if (!string.IsNullOrWhiteSpace(_update.Version))
        {
            SetSpace(10);

            sb.AppendLine($"""
                <div style="font-size:11pt; font-style:italic; color:#666666;">
                    {nameof(TempoOutlookSync)} Version {WebUtility.HtmlEncode(_update.Version)}
                </div>
                """);
        }

        sb.AppendLine("</div>");

        return sb.ToString();

        void SetSpace(int spacingPx) => sb.AppendLine($"""<div style="height:{spacingPx}px; line-height:{spacingPx}px; font-size:{spacingPx}px;">&nbsp;</div>""");
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