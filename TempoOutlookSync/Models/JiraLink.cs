namespace TempoOutlookSync.Models;

using TempoOutlookSync.Dto;

public sealed record JiraLink
{
    public string Id { get; }
    public JiraLinkType Type { get; }
    public string? RelationToBaseIssue { get; }
    public JiraLinkedIssue LinkedIssue { get; }

    public JiraLink(JiraIssueLinkDto dto, string baseUrl)
    {
        Id = dto.Id;
        Type = GetLinkType(dto);
        RelationToBaseIssue = Type switch
        {
            JiraLinkType.Inward => dto.Type.Inward,
            JiraLinkType.Outward => dto.Type.Outward,
            _ => null
        };
        LinkedIssue = new JiraLinkedIssue(Type switch
        {
            JiraLinkType.Inward => dto.InwardIssue ?? throw new InvalidOperationException($"The linked inward issue can not be null"),
            JiraLinkType.Outward => dto.OutwardIssue ?? throw new InvalidOperationException($"The linked outward issue can not be null"),
            _ => throw new InvalidOperationException($"The linked issue can not be {nameof(JiraLinkType.Unknown)}")
        }, baseUrl);
    }

    private static JiraLinkType GetLinkType(JiraIssueLinkDto dto)
    {
        if (dto.InwardIssue is not null) return JiraLinkType.Inward;
        if (dto.OutwardIssue is not null) return JiraLinkType.Outward;
        return JiraLinkType.Unknown;
    }
}