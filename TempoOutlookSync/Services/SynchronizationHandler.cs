namespace TempoOutlookSync.Services;

using System.Net;
using TempoOutlookSync.Common;
using TempoOutlookSync.Models;

public sealed class SynchronizationHandler
{
    private readonly ILogger _logger;
    private readonly ConfigurationHandler _config;
    private readonly TempoApiClient _tempo;
    private readonly JiraApiClient _jira;
    private readonly OutlookComClient _outlook;

    public SynchronizationHandler(ILogger logger, ConfigurationHandler config, TempoApiClient tempo, JiraApiClient jira, OutlookComClient outlook)
    {
        _logger = logger;
        _config = config;
        _tempo = tempo;
        _jira = jira;
        _outlook = outlook;
    }

    public async Task ExecuteAsync()
    {
        try
        {
            _logger.LogInfo("Sync started");

            await _tempo.ThrowIfCantConnect();
            await _jira.ThrowIfCantConnect();

            var today = DateTime.Today.AddDays(-7);
            var todayAddYear = today.AddMonths(6);

            var categoryMappings = await GetCategoryMappingsAsync();

            var existingTempoAppointments = _outlook.GetTempoAppointments()
                .GroupBy(x => x.TempoId)
                .ToDictionary(x => x.Key, x => x.ToHashSet());

            var totalTempoEntries = 0;
            var changeCount = 0;
            await foreach (var entry in _tempo.GetPlannerEntriesAsync(today, todayAddYear))
            {
                totalTempoEntries++;

                var appointmentInfo = await GetAppointmentInfoAsync(entry, categoryMappings);

                var needsCreation = true;
                if (existingTempoAppointments.TryGetValue(entry.Id, out var appointments))
                {
                    needsCreation = false;

                    foreach (var appointment in appointments)
                    {
                        if (appointment.TempoUpdated != appointmentInfo.TempoEntry.LastUpdated || appointment.JiraUpdated != (appointmentInfo.LastUpdated ?? DateTime.MinValue))
                        {
                            _outlook.DeleteByEntryId(appointment.EntryId);
                            needsCreation = true;
                        }
                    }

                    existingTempoAppointments.Remove(entry.Id);
                }
                if (!needsCreation || appointmentInfo.Category is null) continue;

                changeCount++;
                switch (entry.RecurrenceRule)
                {
                    case TempoRecurrenceRule.Never:
                        _outlook.SaveNonRecurring(appointmentInfo);
                        break;

                    case TempoRecurrenceRule.Weekly or TempoRecurrenceRule.BiWeekly:
                        _outlook.SaveWeeklyRecurring(appointmentInfo);
                        break;

                    case TempoRecurrenceRule.Monthly:
                        _outlook.SaveMonthlyRecurrence(appointmentInfo);
                        break;

                    default: changeCount--; break;
                }
            }

            foreach (var deletedAppointments in existingTempoAppointments.Values)
            {
                foreach (var obsoleteAppointment in deletedAppointments)
                {
                    if (obsoleteAppointment.End < today) continue;

                    changeCount++;
                    _outlook.DeleteByEntryId(obsoleteAppointment.EntryId);
                }
            }

            _outlook.PurgeTrashedTempoAppointments();

            _logger.LogInfo($"Synced {changeCount} Outlook item(s) from {totalTempoEntries} Tempo item(s)");
        }
        catch (HttpRequestException ex) when (ex.StatusCode is HttpStatusCode.Unauthorized)
        {
            _logger.LogError("Could not authorize, please check your credentials in the configuration", null);
        }
        catch (Exception ex)
        {
            _logger.LogError("Sync failed", ex);
        }
    }

    private async Task<IReadOnlyDictionary<string, OutlookCategory>> GetCategoryMappingsAsync()
    {
        var mappings = new Dictionary<string, OutlookCategory>();

        foreach (var category in _config.CategorySettings.Categories)
        {
            await foreach (var id in _jira.SearchIssueIdsAsync(category.JQL))
            {
                mappings.TryAdd(id, new OutlookCategory
                {
                    Name = category.Name,
                    Color = (Microsoft.Office.Interop.Outlook.OlCategoryColor)category.Color
                });
            }
        }

        return mappings;
    }

    private async Task<OutlookAppointmentCreationInfo> GetAppointmentInfoAsync(TempoPlannerEntry entry, IReadOnlyDictionary<string, OutlookCategory> categoryMappings)
    {
        var builder = OutlookAppointmentCreationInfoBuilder.FromTempoEntry(entry);

        var jiraUser = await _jira.GetUserByIdAsync(entry.PlannedByJiraUserId);
        if (jiraUser is not null) builder.WithJiraUser(jiraUser);

        switch (entry.PlanItemType)
        {
            case TempoPlanItemType.Issue:
                var jiraIssue = await _jira.GetIssueByIdAsync(entry.PlanItemId);

                if (jiraIssue is not null)
                {
                    var issueBuilder = builder.WithJiraIssue(jiraIssue);
                    if (categoryMappings.TryGetValue(jiraIssue.Id, out var category)) issueBuilder.WithOutlookCategory(category);

                    return issueBuilder.Build();
                }
                break;

            case TempoPlanItemType.Project:
                var jiraProject = await _jira.GetProjectByIdAsync(entry.PlanItemId);

                if (jiraProject is not null)
                {
                    var projectBuilder = builder.WithJiraProject(jiraProject);
                    if (categoryMappings.TryGetValue(jiraProject.Id, out var category)) projectBuilder.WithOutlookCategory(category);

                    return projectBuilder.Build();
                }
                break;
        }

        return builder.Build();
    }
}