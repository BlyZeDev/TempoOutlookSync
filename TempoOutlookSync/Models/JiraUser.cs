namespace TempoOutlookSync.Models;

using TempoOutlookSync.Dto;

public sealed record JiraUser
{
    public string? EmailAddress { get; }
    public string? DisplayName { get; }

    public JiraUser(JiraUserDto dto)
    {
        EmailAddress = dto.EmailAddress;
        DisplayName = dto.DisplayName;
    }
}