using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TempoOutlookSync
{
    sealed class Program
    {
        private const string OutlookPropertyId = "TempoId";
        private static readonly TimeSpan Interval = TimeSpan.FromMinutes(5);

        static async Task Main()
        {
            var handler = ConfigurationHandler.Initialize();
            if (!handler.IsValid)
            {
                Console.Write("Jira User Id: ");
                var userId = Console.ReadLine().Trim();

                Console.Write("Tempo API Key: ");
                var apiToken = Console.ReadLine().Trim();

                handler.SetConfiguration(new Configuration(apiToken, userId));

                Console.Clear();
            }

            if (!handler.IsValid)
            {
                handler.DeleteConfiguration();
                Console.WriteLine("Could not use the configuration");
                Console.ReadLine();
                Environment.Exit(0);
            }

            using (var tempoClient = new TempoClient(handler.Configuration.UserId, handler.Configuration.ApiToken))
            {
                var isConnectionPossible = await tempoClient.CheckIfConnectionPossible();

                if (!isConnectionPossible)
                {
                    handler.DeleteConfiguration();
                    Console.WriteLine("No connection was possible. Double check your credentials");
                    Console.ReadLine();
                    Environment.Exit(0);
                }

                Console.WriteLine("Connection established successfully.");
                Console.WriteLine("This window can now be minimized");

                var outlook = new Application();

                while (true)
                {
                    try
                    {
                        var dateNow = DateTime.Now;
                        var dateEnd = dateNow.AddYears(1);

                        var items = outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items;
                        items.IncludeRecurrences = false;
                        items.Sort("[Start]");
                        items = items.Restrict($"[Start] >= '{dateNow:MM/dd/yyyy HH:mm}' AND [Start] <= '{dateEnd:MM/dd/yyyy HH:mm}' AND [MessageClass] = 'IPM.Appointment'");

                        var existingIds = new HashSet<string>();
                        foreach (AppointmentItem item in items)
                        {
                            var userProp = item.UserProperties.Find(OutlookPropertyId);
                            if (userProp != null) existingIds.Add(userProp.Value);
                        }

                        foreach (var entry in await tempoClient.GetPlannerEntriesAsync(dateNow, dateEnd))
                        {
                            if (existingIds.Contains(entry.Id.ToString())) continue;

                            if (entry.RecurrenceRule is RecurrenceRule.Never && entry.End.Date > entry.Start.Date)
                            {
                                for (var day = entry.Start.Date; day <= entry.End.Date; day = day.AddDays(1))
                                    CreateSingle(outlook, entry, day + entry.StartTime);

                                continue;
                            }

                            if (entry.RecurrenceRule is RecurrenceRule.Monthly && entry.Start.Day != entry.End.Day)
                            {
                                for (var day = entry.Start.Date; day <= entry.End.Date; day = day.AddDays(1))
                                {
                                    var monthlyStart = day + entry.StartTime;

                                    var monthlyAppointment = CreateBase(outlook, entry, monthlyStart);

                                    ApplyRecurrence(monthlyAppointment, entry, monthlyStart);

                                    monthlyAppointment.Save();
                                }

                                continue;
                            }

                            var baseStart = entry.Start.Date + entry.StartTime;
                            var appointment = CreateBase(outlook, entry, baseStart);

                            if (entry.RecurrenceRule is RecurrenceRule.Never) appointment.End = baseStart + entry.DurationPerDay;
                            else ApplyRecurrence(appointment, entry, baseStart);

                            appointment.Save();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    await Task.Delay(Interval);
                }
            }
        }

        private static AppointmentItem CreateBase(Application outlook, TempoPlannerEntry entry, DateTime start)
        {
            var appointment = (AppointmentItem)outlook.CreateItem(OlItemType.olAppointmentItem);

            appointment.Subject = entry.Description;
            appointment.Body = $"[AutoImport by Jira Tempo]\n{entry.Description}";
            appointment.Start = start;
            appointment.BusyStatus = OlBusyStatus.olBusy;
            appointment.ReminderSet = true;
            appointment.ReminderMinutesBeforeStart = 15;

            appointment.UserProperties.Add(OutlookPropertyId, OlUserPropertyType.olText).Value = entry.Id.ToString();

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
                recurrence.DayOfWeekMask = BuildMask(entry.Start.Date, entry.End.Date);
            }

            recurrence.NoEndDate = false;
            
            recurrence.PatternStartDate = entry.Start.Date;
            recurrence.PatternEndDate = entry.RecurrenceEnd.Date;

            recurrence.StartTime = start;
            recurrence.EndTime = start + entry.DurationPerDay;
        }

        private static OlDaysOfWeek BuildMask(DateTime start, DateTime end)
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

            return mask;
        }
    }
}