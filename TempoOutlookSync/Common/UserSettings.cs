namespace TempoOutlookSync.Common;

using CsToml;

[TomlSerializedObject]
public sealed partial record UserSettings
{
    [TomlValueOnSerialized]
    public required string Email { get; init; }
    [TomlValueOnSerialized]
    public required string JiraApiToken { get; init; }
    [TomlValueOnSerialized]
    public required string UserId { get; init; }
    [TomlValueOnSerialized]
    public required string TempoApiToken { get; init; }
}