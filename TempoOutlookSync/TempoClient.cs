using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using System.Threading.Tasks;

namespace TempoOutlookSync
{
    public sealed class TempoClient : IDisposable
    {
        private const string BaseApiUrl = "https://api.tempo.io/4";
        private const string TempoDateFormat = "yyyy-MM-dd";

        private readonly HttpClient _client;
        private readonly string _userId;

        public TempoClient(string userId, string apiToken)
        {
            _userId = userId;

            _client = new HttpClient()
            {
                Timeout = TimeSpan.FromSeconds(30)
            };
            _client.DefaultRequestHeaders.Accept.Clear();
            _client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json")
            {
                CharSet = "utf-8"
            });
            _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiToken);
        }

        public async Task<bool> CheckIfConnectionPossible()
        {
            try
            {
                var response = await _client.GetAsync($"{BaseApiUrl}/plans/user/{_userId}?from={DateTime.Now.ToString(TempoDateFormat)}&to={DateTime.Now.AddDays(1).ToString(TempoDateFormat)}");
                response.EnsureSuccessStatusCode();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }
        }

        public async Task<IEnumerable<TempoPlannerEntry>> GetPlannerEntriesAsync(DateTime startDate, DateTime endDate)
        {
            try
            {
                var response = await _client.GetAsync($"{BaseApiUrl}/plans/user/{_userId}?from={startDate.ToString(TempoDateFormat)}&to={endDate.ToString(TempoDateFormat)}");
                response.EnsureSuccessStatusCode();

                var jsonStream = await response.Content.ReadAsStreamAsync();
                var jsonDoc = await JsonDocument.ParseAsync(jsonStream);
                var root = jsonDoc.RootElement;

                var entries = new List<TempoPlannerEntry>();

                if (root.TryGetProperty("results", out var results))
                {
                    foreach (var result in results.EnumerateArray())
                    {
                        var id = result.GetProperty("id").GetInt32();
                        var hasDescription = result.TryGetProperty("description", out var description);

                        entries.Add(new TempoPlannerEntry(
                            id,
                            DateTime.ParseExact(result.GetProperty("startDate").GetString(), TempoDateFormat, CultureInfo.InvariantCulture),
                            DateTime.ParseExact(result.GetProperty("endDate").GetString(), TempoDateFormat, CultureInfo.InvariantCulture),
                            hasDescription ? description.GetString() : $"Issue #{id}",
                            TimeSpan.ParseExact(result.GetProperty("startTime").GetString(), @"hh\:mm", CultureInfo.InvariantCulture),
                            TimeSpan.FromSeconds(result.GetProperty("plannedSecondsPerDay").GetInt64()),
                            ParseRecurrenceRule(result.GetProperty("rule").GetString()),
                            DateTime.ParseExact(result.GetProperty("recurrenceEndDate").GetString(), TempoDateFormat, CultureInfo.InvariantCulture)));
                    }
                }

                return entries;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return Enumerable.Empty<TempoPlannerEntry>();
            }
        }

        public void Dispose()
        {
            _client.Dispose();
        }

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
}
