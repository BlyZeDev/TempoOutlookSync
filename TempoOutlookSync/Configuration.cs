using System;
using System.Text.Json.Serialization;

namespace TempoOutlookSync
{
    [Serializable]
    public sealed class Configuration
    {
        [JsonInclude]
        public string ApiToken { get; }
        [JsonInclude]
        public string UserId { get; }

        [JsonConstructor]
        public Configuration(string apiToken, string userId)
        {
            ApiToken = apiToken;
            UserId = userId;
        }
    }
}
