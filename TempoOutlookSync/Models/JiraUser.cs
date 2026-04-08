namespace TempoOutlookSync.Models;

using TempoOutlookSync.Dto;

public sealed record JiraUser
{
    public string? EmailAddress { get; }
    public string? DisplayName { get; }
    public string? AvatarUrl { get; }

    public JiraUser(JiraUserDto dto)
    {
        EmailAddress = dto.EmailAddress;
        DisplayName = dto.DisplayName;
        AvatarUrl = dto.AvatarUrls?.Avatar48 ?? dto.AvatarUrls?.Avatar32 ?? dto.AvatarUrls?.Avatar24 ?? dto.AvatarUrls?.Avatar16 ?? null;
    }
}