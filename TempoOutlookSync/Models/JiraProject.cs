namespace TempoOutlookSync.Models;

using TempoOutlookSync.Dto;

public sealed record JiraProject
{
    public string Id { get; }
    public string Key { get; }
    public string Permalink { get; }
    public string? Name { get; }

    public JiraProject(JiraProjectDto dto, string baseUrl)
    {
        Id = dto.Id;
        Key = dto.Key;
        Permalink = $"{baseUrl}{Key}";
        Name = dto.Name;
    }
}