namespace TempoOutlookSync.Services;

using System.Globalization;
using System.Net.Http.Headers;
using System.Text.Json;
using TempoOutlookSync.Common;
using TempoOutlookSync.Dto;

public sealed class TempoApiClient : IDisposable
{
    private const string BaseApiUrl = "https://api.tempo.io/4";
    public const string TempoDateFormat = "yyyy-MM-dd";

    private readonly ILogger _logger;
    private readonly ConfigurationHandler _config;

    private readonly HttpClient _client;

    public TempoApiClient(ILogger logger, ConfigurationHandler config)
    {
        _logger = logger;
        _config = config;

        _client = new HttpClient()
        {
            Timeout = TimeSpan.FromSeconds(30)
        };
    }

    public async Task ThrowIfCantConnect()
    {
        SetHeaders(_client, _config.Current);
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
            SetHeaders(_client, _config.Current);
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
            _logger.LogError(ex.Message, ex);
            yield break;
        }

        foreach (var entry in payload.Results ?? [])
        {
            yield return new TempoPlannerEntry(entry);
        }
    }

    public void Dispose() => _client.Dispose();

    private static void SetHeaders(HttpClient client, Configuration config)
    {
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json")
        {
            CharSet = "utf-8"
        });
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", config.TempoApiToken);
    }

    private static string BuildTempoPlannerUrl(DateTime startDate, DateTime endDate, Configuration config)
        => $"{BaseApiUrl}/plans/user/{config.UserId}?from={startDate.ToString(TempoDateFormat)}&to={endDate.ToString(TempoDateFormat)}";
}
