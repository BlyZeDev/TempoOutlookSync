namespace TempoOutlookSync.Services;

using System.Buffers.Text;
using System.Net.Http.Headers;
using System.Text;
using TempoOutlookSync.Common;

public sealed class JiraApiClient : IDisposable
{
    private const string BaseApiUrl = "https://edocag.atlassian.net/rest/api/3";

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

    //GetIssueById

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