namespace TempoOutlookSync.Services;

using DotTray;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Net;
using TempoOutlookSync.Common;
using TempoOutlookSync.Models;

public sealed class ServiceRunner : IDisposable
{
    private static readonly TimeSpan Interval = TimeSpan.FromMinutes(15);

    private readonly ILogger _logger;
    private readonly TempoOutlookSyncContext _context;
    private readonly UpdateHandler _update;
    private readonly ConfigurationHandler _config;
    private readonly TempoApiClient _tempo;
    private readonly JiraApiClient _jira;
    private readonly OutlookComClient _outlook;

    private readonly CancellationTokenSource _cts;
    private readonly NotifyIcon _icon;
    private readonly MenuItem _syncNowMenuItem;
    private readonly MenuItem _nextSyncMenuItem;

    private CancellationTokenSource manualSyncCts;
    private long lastSyncUtcBinary;
    private bool isSyncing;

    public ServiceRunner(ILogger logger, TempoOutlookSyncContext context, UpdateHandler update, ConfigurationHandler config, TempoApiClient tempo, JiraApiClient jira, OutlookComClient outlook)
    {
        _logger = logger;
        _context = context;
        _update = update;
        _config = config;
        _tempo = tempo;
        _jira = jira;
        _outlook = outlook;

        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnhandledTaskException;

        _cts = new CancellationTokenSource();
        _icon = NotifyIcon.Run(_context.IcoPath, _cts.Token, x =>
        {
            x.BackgroundHoverColor = new TrayColor(218, 83, 225);
            x.BackgroundDisabledColor = new TrayColor(40, 40, 40);
            x.TextDisabledColor = new TrayColor(180, 180, 180);
        }, x => x.LineThickness = 1.2f);
        _icon.SetFontSize(16f);
        _icon.SetToolTip($"{nameof(TempoOutlookSync)} - Version {_update.Version}");
        _icon.PopupShowing += OnPopupShowing;

        manualSyncCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token);

        _syncNowMenuItem = _icon.MenuItems.AddItem(x =>
        {
            x.Text = "Sync now";
            x.Clicked = _ => manualSyncCts.Cancel();
        });
        _nextSyncMenuItem = _icon.MenuItems.AddItem(x =>
        {
            x.Text = $"Next Sync in {Util.FormatTime(Interval)}";
            x.IsDisabled = true;
        });
        _icon.MenuItems.AddItem(x =>
        {
            x.Text = "Settings";
            x.SubMenu.AddItem(x =>
            {
                x.Text = "Open Application Folder";
                x.Clicked = _ =>
                {
                    _logger.LogDebug("Opening the application folder");
                    Directory.CreateDirectory(context.AppFilesDirectory);
                    Util.ShellOpen(_context.AppFilesDirectory);
                };
            });
            x.SubMenu.AddItem(x =>
            {
                x.Text = "Autostart";
                x.IsChecked = Util.IsInStartup(nameof(TempoOutlookSync), _context.ExecutablePath);
                x.Clicked = args =>
                {
                    Util.RemoveFromStartup(nameof(TempoOutlookSync));
                    if (args.MenuItem.IsChecked.GetValueOrDefault()) Util.AddToStartup(nameof(TempoOutlookSync), _context.ExecutablePath);

                    var isActivated = Util.IsInStartup(nameof(TempoOutlookSync), _context.ExecutablePath);
                    args.MenuItem.IsChecked = isActivated;

                    _logger.LogInfo($"Autostart is now {(isActivated ? "activated" : "removed")}");
                };
            });
            x.SubMenu.AddItem(x =>
            {
                x.Text = "Help";
                x.Clicked = _ => Util.ShellOpen(_context.GitHubRepoUrl);
            });
        });
        _icon.MenuItems.AddSeparator();
        _icon.MenuItems.AddItem(x =>
        {
            x.Text = $"Version - {_update.Version}";
            x.TextDisabledColor = x.TextColor;
            x.IsDisabled = true;
        });
        _icon.MenuItems.AddSeparator();
        _icon.MenuItems.AddItem(x =>
        {
            x.Text = "Exit";
            x.Clicked = _ => _cts.Cancel();
        });
    }

    public async Task RunAsync()
    {
        _logger.LogLevel = LogLevel.Debug;
        _logger.Log += OnLog;
        _config.UserSettingsChanged += UserSettingsChanged;

        _icon.ShowBalloon(new BalloonNotification
        {
            Icon = BalloonNotificationIcon.User,
            Title = $"{nameof(TempoOutlookSync)}",
            Message = "is now running in the background"
        });

        while (!_cts.IsCancellationRequested)
        {
            manualSyncCts.Dispose();
            manualSyncCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token);

            await SyncTempoToOutlookAsync();

            using (var process = Process.GetCurrentProcess())
            {
                process.Refresh();
                _logger.LogDebug($"Memory Allocated: {Util.FormatBytes(process.PrivateMemorySize64)}");
            }

            try
            {
                await Task.Delay(Interval, manualSyncCts.Token);
            }
            catch (OperationCanceledException) { }
        }
    }

    public void Dispose()
    {
        _icon.Dispose();
        manualSyncCts.Dispose();
        _cts.Dispose();
    }

    private void OnPopupShowing(MouseButton mouseButton)
    {
        _nextSyncMenuItem.Text = Interlocked.CompareExchange(ref isSyncing, false, false)
            ? $"Syncing..."
            : $"Next Sync in {Util.FormatTime(GetRemainingUntilSync(lastSyncUtcBinary))}";
    }

    private void OnLog(LogLevel logLevel, string message, Exception? exception)
    {
        if (logLevel < LogLevel.Error) return;

        _icon.ShowBalloon(new BalloonNotification
        {
            Icon = BalloonNotificationIcon.User,
            Title = logLevel.ToString(),
            Message = exception?.Message ?? message
        });
    }

    private void UserSettingsChanged(ObjectChangedEventArgs<UserSettings> args)
    {
        if (!args.Old.UserId.Equals(args.New.UserId, StringComparison.Ordinal)
            || !args.Old.TempoApiToken.Equals(args.New.TempoApiToken, StringComparison.Ordinal)
            || !args.Old.Email.Equals(args.New.Email, StringComparison.Ordinal)
            || !args.Old.JiraApiToken.Equals(args.New.JiraApiToken, StringComparison.Ordinal)) manualSyncCts.Cancel();
    }

    private async Task SyncTempoToOutlookAsync()
    {
        if (Interlocked.Exchange(ref isSyncing, true)) return;
        _syncNowMenuItem.IsDisabled = true;

        try
        {
            await _tempo.ThrowIfCantConnect();
            await _jira.ThrowIfCantConnect();

            _logger.LogInfo("Sync started");

            var today = DateTime.Today.AddDays(-7);
            var todayAddYear = today.AddYears(1);

            var existingTempoAppointments = _outlook.GetOutlookTempoAppointments(today);

            var changeCount = 0;
            await foreach (var entry in _tempo.GetPlannerEntriesAsync(today, todayAddYear))
            {
                var appointmentInfo = await GetAppointmentInfoAsync(entry);

                var needsCreation = true;
                if (existingTempoAppointments.TryGetValue(entry.Id, out var appointments))
                {
                    needsCreation = false;

                    foreach (var item in appointments)
                    {
                        switch (appointmentInfo)
                        {
                            case null: _outlook.DeleteByEntryId(item.EntryId); break;

                            case var _ when item.TempoUpdated != entry.LastUpdated || item.JiraUpdated != (appointmentInfo.LastUpdated ?? DateTime.MinValue):
                                _outlook.DeleteByEntryId(item.EntryId);
                                needsCreation = true;
                                break;
                        }
                    }

                    existingTempoAppointments.Remove(entry.Id);
                }
                if (!needsCreation || appointmentInfo is null) continue;

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

            _logger.LogInfo(@$"Synced {changeCount} item(s), next sync in {Util.FormatTime(Interval)}");
        }
        catch (HttpRequestException ex) when (ex.StatusCode is HttpStatusCode.Unauthorized)
        {
            _logger.LogError("Could not authorize, please check your credentials in the configuration", null);
        }
        catch (Exception ex)
        {
            _logger.LogError("Sync failed", ex);
        }
        finally
        {
            Interlocked.Exchange(ref lastSyncUtcBinary, DateTime.UtcNow.ToBinary());
            Interlocked.Exchange(ref isSyncing, false);
            _syncNowMenuItem.IsDisabled = false;
        }
    }

    private async Task<OutlookAppointmentInfo?> GetAppointmentInfoAsync(TempoPlannerEntry entry)
    {
        OutlookAppointmentInfo appointmentInfo;
        switch (entry.PlanItemType)
        {
            case TempoPlanItemType.Issue:
                var jiraIssue = await _jira.GetIssueByIdAsync(entry.PlanItemId);
                if (jiraIssue?.StatusCategory is JiraStatusCategory.Done) return null;
                appointmentInfo = jiraIssue is null ? new OutlookAppointmentInfo(entry) : new OutlookAppointmentInfo(entry, jiraIssue);
                break;

            case TempoPlanItemType.Project:
                var jiraProject = await _jira.GetProjectByIdAsync(entry.PlanItemId);
                appointmentInfo = jiraProject is null ? new OutlookAppointmentInfo(entry) : new OutlookAppointmentInfo(entry, jiraProject);
                break;

            default: appointmentInfo = new OutlookAppointmentInfo(entry); break;
        }
        return appointmentInfo;
    }


    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs args)
        => ControlledCrash(args.ExceptionObject as Exception ?? new Exception("Unknown exception was thrown"));

    private void OnUnhandledTaskException(object? sender, UnobservedTaskExceptionEventArgs args)
    {
        args.SetObserved();
        ControlledCrash(args.Exception);
    }

    [DoesNotReturn]
    private void ControlledCrash(Exception exception)
    {
        _logger.LogCritical("The application crashed", exception);

        var crashLogPath = _context.WriteCrashLog(exception);
        Environment.FailFast(exception.Message, exception);
    }

    private static TimeSpan GetRemainingUntilSync(long lastSyncUtcBinary)
    {
        var remaining = Interval - (DateTime.UtcNow - DateTime.FromBinary(lastSyncUtcBinary));
        return remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero;
    }
}