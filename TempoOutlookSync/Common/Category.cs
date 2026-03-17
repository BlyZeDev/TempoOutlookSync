namespace TempoOutlookSync.Common;

using CsToml;

[TomlSerializedObject]
public sealed partial record Category
{
    [TomlValueOnSerialized]
    public required string Name { get; init; }
    [TomlValueOnSerialized]
    public required OutlookColor Color { get; init; }
    [TomlValueOnSerialized]
    public required string JQL { get; init; }
}