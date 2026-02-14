using System.Text.Json;

namespace TempoOutlookSync.Models;

public sealed class JiraProject
{
    private static readonly JsonSerializerOptions _options = new JsonSerializerOptions()
    {
        WriteIndented = true
    };

    public override string ToString() => JsonSerializer.Serialize(this, _options);
}