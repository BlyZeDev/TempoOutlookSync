namespace TempoOutlookSync.Common;

using CsToml;

[TomlSerializedObject]
public sealed partial record CategorySettings
{
    [TomlValueOnSerialized]
    public IEnumerable<Category> Categories { get; init; } = [];
}