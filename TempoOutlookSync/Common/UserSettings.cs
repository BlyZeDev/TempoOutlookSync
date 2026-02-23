namespace TempoOutlookSync.Common;

using CsToml;

[TomlSerializedObject]
public sealed partial record UserSettings
{
    [TomlValueOnSerialized]
    public string Email { get; init; } = "";
    [TomlValueOnSerialized]
    public string JiraApiToken { get; init; } = "";
    [TomlValueOnSerialized]
    public string UserId { get; init; } = "";
    [TomlValueOnSerialized]
    public string TempoApiToken { get; init; } = "";
}