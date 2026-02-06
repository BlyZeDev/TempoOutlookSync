using System;
using System.Linq;
using System.Threading.Tasks;

namespace TempoOutlookSync
{
    sealed class Program
    {
        private static readonly TimeSpan Interval = TimeSpan.FromMinutes(15);

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
                Console.WriteLine();

                using (var outlookClient = new OutlookClient())
                {
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("Sync started");

                            var today = DateTime.Today;
                            var todayAddYear = today.AddYears(1);

                            var existingTempoAppointments = outlookClient.GetOutlookTempoAppointments(today);

                            var changeCount = 0;
                            foreach (var entry in await tempoClient.GetPlannerEntriesAsync(today, todayAddYear))
                            {
                                var entryId = entry.Id.ToString();
                                var needsCreation = true;

                                if (existingTempoAppointments.TryGetValue(entryId, out var appointments))
                                {
                                    needsCreation = appointments.Any(item => outlookClient.DeleteIfOutdated(item, entry.LastUpdated));
                                    existingTempoAppointments.Remove(entryId);
                                }

                                if (!needsCreation) continue;

                                changeCount++;
                                switch (entry.RecurrenceRule)
                                {
                                    case RecurrenceRule.Never when entry.End.Date >= entry.Start.Date:
                                        outlookClient.SaveNonRecurring(entry);
                                        break;

                                    case RecurrenceRule.Weekly:
                                    case RecurrenceRule.BiWeekly:
                                        outlookClient.SaveWeeklyRecurring(entry);
                                        break;

                                    case RecurrenceRule.Monthly when entry.End.Day != entry.Start.Day:
                                        outlookClient.SaveMonthlyRecurrence(entry);
                                        break;

                                    default: changeCount--; break;
                                }
                            }

                            foreach (var deletedAppointments in existingTempoAppointments.Values)
                            {
                                foreach (var obsoleteAppointment in deletedAppointments)
                                {
                                    if (obsoleteAppointment.End < today) continue;

                                    changeCount++;
                                    obsoleteAppointment.Delete();
                                }
                            }

                            Console.WriteLine($"Synced {changeCount} item{(changeCount == 1 ? char.MinValue : 's')}, next sync in {Interval.TotalMinutes:F2} minutes");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                        }

                        await Task.Delay(Interval);
                    }
                }
            }
        }
    }
}