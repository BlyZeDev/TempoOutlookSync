namespace TempoOutlookSync.Services;

using System.Buffers.Text;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using TempoOutlookSync.Common;
using TempoOutlookSync.Dto;
using TempoOutlookSync.Models;

public sealed class JiraApiClient : IDisposable
{
    private const string BaseUrl = "https://edocag.atlassian.net";
    private const string BaseApiUrl = $"{BaseUrl}/rest/api/3";

    private readonly ILogger _logger;
    private readonly ConfigurationHandler _config;

    private readonly HttpClient _client;

    public JiraApiClient(ILogger logger, ConfigurationHandler config)
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

        var url = $"{BaseApiUrl}/myself";
        using (var response = await _client.GetAsync(url))
        {
            response.EnsureSuccessStatusCode();
        }
    }

    public async Task<JiraIssue?> GetIssueByIdAsync(string? id)
    {
        if (id is null) return null;

        try
        {
            SetHeaders(_client, _config.Current);

            var url = $"{BaseApiUrl}/issue/{id}";
            using (var response = await _client.GetAsync(url))
            {
                response.EnsureSuccessStatusCode();

                using (var stream = await response.Content.ReadAsStreamAsync())
                {
                    var issueDto = await JsonSerializer.DeserializeAsync<JiraIssueDto>(stream, JiraIssueDtoJsonContext.Default.JiraIssueDto);

                    return issueDto is null ? null : new JiraIssue(issueDto, $"{BaseUrl}/browse/");
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex.Message, ex);
            return null;
        }
    }

    public async Task<JiraProject?> GetProjectByIdAsync(string? id)
    {
        if (id is null) return null;
    }

    public void Dispose() => _client.Dispose();

    private static void SetHeaders(HttpClient client, Configuration config)
    {
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json")
        {
            CharSet = "utf-8"
        });

        var base64 = Encoding.UTF8.GetBytes($"{config.Email}:{config.JiraApiToken}");
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(base64));
    }
}