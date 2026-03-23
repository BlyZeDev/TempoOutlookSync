namespace TempoOutlookSync.Common;

using TempoOutlookSync.Models;

public sealed class OutlookAppointmentInfoBuilder : IOutlookAppointmentInfoBuilder, IWithJiraIssueSelected, IWithJiraProjectSelected
{
    private readonly TempoPlannerEntry tempoPlannerEntry;
    private JiraUser? jiraUser;
    private OutlookCategory? outlookCategory;
    private JiraIssue? jiraIssue;
    private JiraProject? _jiraProject;

    private OutlookAppointmentInfoBuilder(TempoPlannerEntry tempoPlannerEntry) => this.tempoPlannerEntry = tempoPlannerEntry;

    IOutlookAppointmentInfoBuilder IOutlookAppointmentInfoBuilder.WithJiraUser(JiraUser jiraUser) => SetJiraUser(jiraUser);
    IWithJiraIssueSelected IWithJiraIssueSelected.WithJiraUser(JiraUser jiraUser) => SetJiraUser(jiraUser);
    IWithJiraProjectSelected IWithJiraProjectSelected.WithJiraUser(JiraUser jiraUser) => SetJiraUser(jiraUser);

    IWithJiraIssueSelected IWithJiraIssueSelected.WithOutlookCategory(OutlookCategory outlookCategory) => SetOutlookCategory(outlookCategory);
    IWithJiraProjectSelected IWithJiraProjectSelected.WithOutlookCategory(OutlookCategory outlookCategory) => SetOutlookCategory(outlookCategory);
    
    public IWithJiraIssueSelected WithJiraIssue(JiraIssue jiraIssue)
    {
        this.jiraIssue = jiraIssue;
        return this;
    }

    public IWithJiraProjectSelected WithJiraProject(JiraProject jiraProject)
    {
        _jiraProject = jiraProject;
        return this;
    }

    public OutlookAppointmentInfo Build()
    {
        if (jiraIssue is not null) return new OutlookAppointmentInfo(tempoPlannerEntry, jiraIssue, jiraUser, outlookCategory);
        if (_jiraProject is not null) return new OutlookAppointmentInfo(tempoPlannerEntry, _jiraProject, jiraUser, outlookCategory);

        return new OutlookAppointmentInfo(tempoPlannerEntry, jiraUser);
    }

    private OutlookAppointmentInfoBuilder SetJiraUser(JiraUser jiraUser)
    {
        this.jiraUser = jiraUser;
        return this;
    }

    private OutlookAppointmentInfoBuilder SetOutlookCategory(OutlookCategory outlookCategory)
    {
        this.outlookCategory = outlookCategory;
        return this;
    }

    public static IOutlookAppointmentInfoBuilder FromTempoEntry(TempoPlannerEntry tempoPlannerEntry) => new OutlookAppointmentInfoBuilder(tempoPlannerEntry);
}

public interface IOutlookAppointmentInfoBuilder
{
    public IOutlookAppointmentInfoBuilder WithJiraUser(JiraUser jiraUser);
    public IWithJiraIssueSelected WithJiraIssue(JiraIssue jiraIssue);
    public IWithJiraProjectSelected WithJiraProject(JiraProject jiraProject);
    public OutlookAppointmentInfo Build();
}

public interface IWithJiraIssueSelected
{
    public IWithJiraIssueSelected WithJiraUser(JiraUser jiraUser);
    public IWithJiraIssueSelected WithOutlookCategory(OutlookCategory outlookCategory);
    public OutlookAppointmentInfo Build();
}

public interface IWithJiraProjectSelected
{
    public IWithJiraProjectSelected WithJiraUser(JiraUser jiraUser);
    public IWithJiraProjectSelected WithOutlookCategory(OutlookCategory outlookCategory);
    public OutlookAppointmentInfo Build();
}