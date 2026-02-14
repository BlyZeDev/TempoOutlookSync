namespace TempoOutlookSync.Models;

using Microsoft.Office.Interop.Outlook;

public sealed record OutlookCategory
{
    public required string Name { get; init; }
    public required OlCategoryColor Color { get; init; }
}