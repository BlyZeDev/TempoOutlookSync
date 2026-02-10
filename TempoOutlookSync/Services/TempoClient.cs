namespace TempoOutlookSync.Services;

using System.Globalization;
using System.Net.Http.Headers;
using System.Text.Json;
using TempoOutlookSync.Common;

public sealed class TempoClient : IDisposable
{
    private const string BaseApiUrl = "https://api.tempo.io/4";
    public const string TempoDateFormat = "yyyy-MM-dd";

    private readonly ConfigurationHandler _config;

    private readonly HttpClient _client;

    public TempoClient(ConfigurationHandler config)
    {
        _config = config;

        _client = new HttpClient()
        {
            Timeout = TimeSpan.FromSeconds(30)
        };
    }

    public async Task ThrowIfCantConnect()
    {
        using (var response = await _client.GetAsync(BuildTempoPlannerUrl(DateTime.Now, DateTime.Now.AddDays(1), _config.Current)))
        {
            response.EnsureSuccessStatusCode();
        }
    }

    public async IAsyncEnumerable<TempoPlannerEntry> GetPlannerEntriesAsync(DateTime startDate, DateTime endDate)
    {
        TempoPlannerPayloadDto payload;

        try
        {
            using (var response = await _client.GetAsync(BuildTempoPlannerUrl(startDate, endDate, _config.Current)))
            {
                response.EnsureSuccessStatusCode();

                using (var stream = await response.Content.ReadAsStreamAsync())
                {
                    payload = await JsonSerializer.DeserializeAsync<TempoPlannerPayloadDto>(stream) ?? throw new JsonException("Invalid response");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            yield break;
        }

        foreach (var entry in payload.Results ?? [])
        {
            yield return new TempoPlannerEntry(entry);
        }
    }

    public void Dispose() => _client.Dispose();

    private static void SetRequiredHeaders(HttpClient client, Configuration config)
    {
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json")
        {
            CharSet = "utf-8"
        });
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", config.ApiToken);
    }

    private static string BuildTempoPlannerUrl(DateTime startDate, DateTime endDate, Configuration config)
        => $"{BaseApiUrl}/plans/user/{config.UserId}?from={startDate.ToString(TempoDateFormat)}&to={endDate.ToString(TempoDateFormat)}";
}
