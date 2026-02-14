namespace TempoOutlookSync.Models;

using System.Text.Json;
using TempoOutlookSync.Dto;

public sealed class JiraProject
{
    private static readonly JsonSerializerOptions _options = new JsonSerializerOptions()
    {
        WriteIndented = true
    };

    public string Id { get; }
    public string Key { get; }
    public string Permalink { get; }
    public string? Name { get; }
    public string? Category { get; }

    public JiraProject(JiraProjectDto dto, string baseUrl)
    {
        Id = dto.Id;
        Key = dto.Key;
        Permalink = $"{baseUrl}{Key}";
        Name = dto.Name;
        Category = dto.Category?.Name;
    }

    public override string ToString() => JsonSerializer.Serialize(this, _options);
}