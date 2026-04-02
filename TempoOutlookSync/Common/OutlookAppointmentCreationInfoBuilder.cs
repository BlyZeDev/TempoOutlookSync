namespace TempoOutlookSync.Common;

using TempoOutlookSync.Models;

public sealed class OutlookAppointmentCreationInfoBuilder : IOutlookAppointmentCreationInfoBuilder, IWithJiraIssueSelected, IWithJiraProjectSelected
{
    private readonly TempoPlannerEntry tempoPlannerEntry;
    private JiraUser? jiraUser;
    private OutlookCategory? outlookCategory;
    private JiraIssue? jiraIssue;
    private JiraProject? _jiraProject;

    private OutlookAppointmentCreationInfoBuilder(TempoPlannerEntry tempoPlannerEntry) => this.tempoPlannerEntry = tempoPlannerEntry;

    IOutlookAppointmentCreationInfoBuilder IOutlookAppointmentCreationInfoBuilder.WithJiraUser(JiraUser jiraUser) => SetJiraUser(jiraUser);
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

    public OutlookAppointmentCreationInfo Build()
    {
        if (jiraIssue is not null) return new OutlookAppointmentCreationInfo(tempoPlannerEntry, jiraIssue, jiraUser, outlookCategory);
        if (_jiraProject is not null) return new OutlookAppointmentCreationInfo(tempoPlannerEntry, _jiraProject, jiraUser, outlookCategory);

        return new OutlookAppointmentCreationInfo(tempoPlannerEntry, jiraUser);
    }

    private OutlookAppointmentCreationInfoBuilder SetJiraUser(JiraUser jiraUser)
    {
        this.jiraUser = jiraUser;
        return this;
    }

    private OutlookAppointmentCreationInfoBuilder SetOutlookCategory(OutlookCategory outlookCategory)
    {
        this.outlookCategory = outlookCategory;
        return this;
    }

    public static IOutlookAppointmentCreationInfoBuilder FromTempoEntry(TempoPlannerEntry tempoPlannerEntry) => new OutlookAppointmentCreationInfoBuilder(tempoPlannerEntry);
}

public interface IOutlookAppointmentCreationInfoBuilder
{
    public IOutlookAppointmentCreationInfoBuilder WithJiraUser(JiraUser jiraUser);
    public IWithJiraIssueSelected WithJiraIssue(JiraIssue jiraIssue);
    public IWithJiraProjectSelected WithJiraProject(JiraProject jiraProject);
    public OutlookAppointmentCreationInfo Build();
}

public interface IWithJiraIssueSelected
{
    public IWithJiraIssueSelected WithJiraUser(JiraUser jiraUser);
    public IWithJiraIssueSelected WithOutlookCategory(OutlookCategory outlookCategory);
    public OutlookAppointmentCreationInfo Build();
}

public interface IWithJiraProjectSelected
{
    public IWithJiraProjectSelected WithJiraUser(JiraUser jiraUser);
    public IWithJiraProjectSelected WithOutlookCategory(OutlookCategory outlookCategory);
    public OutlookAppointmentCreationInfo Build();
}