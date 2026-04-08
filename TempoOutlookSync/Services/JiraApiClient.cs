namespace TempoOutlookSync.Services;

using System.Net.Http.Headers;
using System.Net.Http.Json;
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
        SetHeaders(_client, _config.UserSettings);

        var url = $"{BaseApiUrl}/myself";
        using (var response = await _client.GetAsync(url))
        {
            response.EnsureSuccessStatusCode();
        }
    }

    public async Task<JiraUser?> GetUserByIdAsync(string? accountId)
    {
        if (string.IsNullOrWhiteSpace(accountId)) return null;

        try
        {
            SetHeaders(_client, _config.UserSettings);

            var url = $"{BaseApiUrl}/user?accountId={accountId}";
            using (var response = await _client.GetAsync(url))
            {
                response.EnsureSuccessStatusCode();

                using (var stream = await response.Content.ReadAsStreamAsync())
                {
                    var userDto = await JsonSerializer.DeserializeAsync(stream, JiraUserDtoJsonContext.Default.JiraUserDto);

                    return userDto is null ? null : new JiraUser(userDto);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex.Message, ex);
            return null;
        }
    }

    public async IAsyncEnumerable<string> SearchIssueIdsAsync(string? jql)
    {
        if (string.IsNullOrWhiteSpace(jql)) yield break;

        JiraJqlSearchPayloadDto? payload = null;
        do
        {
            try
            {
                SetHeaders(_client, _config.UserSettings);

                var request = new
                {
                    maxResults = 100,
                    nextPageToken = payload?.NextPageToken,
                    jql
                };

                using (var content = JsonContent.Create(request))
                {
                    var url = $"{BaseApiUrl}/search/jql";
                    using (var response = await _client.PostAsync(url, content))
                    {
                        response.EnsureSuccessStatusCode();

                        using (var stream = await response.Content.ReadAsStreamAsync())
                        {
                            payload = await JsonSerializer.DeserializeAsync(stream,
                                JiraJqlSearchPayloadDtoJsonContext.Default.JiraJqlSearchPayloadDto) ?? throw new JsonException("Invalid response");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message, ex);
                yield break;
            }

            foreach (var issue in payload.Issues ?? [])
            {
                yield return issue.Id;
            }
        } while (payload.NextPageToken is not null);
    }

    public async Task<JiraIssue?> GetIssueByIdAsync(string? id)
    {
        if (string.IsNullOrWhiteSpace(id)) return null;

        try
        {
            SetHeaders(_client, _config.UserSettings);

            var url = $"{BaseApiUrl}/issue/{id}?fields=id,key,summary,project,updated,created";
            using (var response = await _client.GetAsync(url))
            {
                response.EnsureSuccessStatusCode();

                using (var stream = await response.Content.ReadAsStreamAsync())
                {
                    var issueDto = await JsonSerializer.DeserializeAsync(stream, JiraIssueDtoJsonContext.Default.JiraIssueDto);

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
        if (string.IsNullOrWhiteSpace(id)) return null;

        try
        {
            SetHeaders(_client, _config.UserSettings);

            var url = $"{BaseApiUrl}/project/{id}";
            using (var response = await _client.GetAsync(url))
            {
                response.EnsureSuccessStatusCode();

                using (var stream = await response.Content.ReadAsStreamAsync())
                {
                    var projectDto = await JsonSerializer.DeserializeAsync(stream, JiraProjectDtoJsonContext.Default.JiraProjectDto);

                    return projectDto is null ? null : new JiraProject(projectDto, $"{BaseUrl}/browse/");
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex.Message, ex);
            return null;
        }
    }

    public void Dispose() => _client.Dispose();

    private static void SetHeaders(HttpClient client, UserSettings settings)
    {
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json")
        {
            CharSet = "utf-8"
        });

        var base64 = Encoding.UTF8.GetBytes($"{settings.Email}:{settings.JiraApiToken}");
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(base64));
    }
}