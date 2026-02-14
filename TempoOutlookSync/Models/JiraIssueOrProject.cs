namespace TempoOutlookSync.Models;

using OneOf;

[GenerateOneOf]
public sealed partial class JiraIssueOrProject : OneOfBase<JiraIssue, JiraProject>;