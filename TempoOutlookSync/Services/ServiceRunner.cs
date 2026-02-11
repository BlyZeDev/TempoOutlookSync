namespace TempoOutlookSync.Services;

using DotTray;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Net;
using System.Runtime.InteropServices;
using TempoOutlookSync.Common;
using TempoOutlookSync.NATIVE;

public sealed class ServiceRunner : IDisposable
{
    private static readonly TimeSpan Interval = TimeSpan.FromMinutes(15);

    private readonly ILogger _logger;
    private readonly TempoOutlookSyncContext _context;
    private readonly ConfigurationHandler _config;
    private readonly TempoClient _tempo;
    private readonly OutlookClient _outlook;

    private readonly CancellationTokenSource _cts;
    private readonly NotifyIcon _icon;
    private readonly MenuItem _nextSyncMenuItem;

    private bool isSyncing;

    public ServiceRunner(ILogger logger, TempoOutlookSyncContext context, ConfigurationHandler config, TempoClient tempo, OutlookClient outlook)
    {
        _logger = logger;
        _context = context;
        _config = config;
        _tempo = tempo;
        _outlook = outlook;

        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnhandledTaskException;

        var hWnd = PInvoke.GetConsoleWindow();
        PInvoke.PostMessage(hWnd, PInvoke.WM_SETICON, PInvoke.ICON_BIG, _context.IcoHandle);
        if (!Program.IsDebug) PInvoke.ShowWindowAsync(hWnd, PInvoke.SW_HIDE);

        _cts = new CancellationTokenSource();
        _icon = NotifyIcon.Run(_context.IcoHandle, _cts.Token, x =>
        {
            x.BackgroundHoverColor = new TrayColor(218, 83, 225);
            x.BackgroundDisabledColor = new TrayColor(40, 40, 40);
            x.TextDisabledColor = new TrayColor(180, 180, 180);
        }, x =>
        {
            x.LineThickness = 1.2f;
        });
        _icon.FontSize = 18f;
        _icon.SetToolTip($"{nameof(TempoOutlookSync)} - Version {TempoOutlookSyncContext.Version}");

        _icon.MenuItems.AddItem(x =>
        {
            x.Text = "Sync now";
            x.Clicked = async _ => await PerformManualSyncAsync();
        });
        _nextSyncMenuItem = _icon.MenuItems.AddItem(x =>
        {
            x.Text = "Next Sync: ";
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
                    ShellOpen(_context.ApplicationDirectory);
                };
            });
            x.SubMenu.AddItem(x =>
            {
                x.Text = "Edit Configuration";
                x.Clicked = _ =>
                {
                    _logger.LogDebug("Opening the configuration file");
                    ShellOpen(_context.ConfigurationPath);
                };
            });
            x.SubMenu.AddItem(x =>
            {
                x.Text = "Autostart";
                x.IsChecked = Util.IsInStartup(nameof(TempoOutlookSync), _context.ExecutablePath);
                x.Clicked = args =>
                {
                    if (Util.IsInStartup(nameof(TempoOutlookSync), _context.ExecutablePath)) Util.RemoveFromStartup(nameof(TempoOutlookSync));
                    else Util.AddToStartup(nameof(TempoOutlookSync), _context.ExecutablePath);

                    var isActivated = Util.IsInStartup(nameof(TempoOutlookSync), _context.ExecutablePath);
                    args.MenuItem.IsChecked = isActivated;

                    _logger.LogInfo($"Autostart is now {(isActivated ? "activated" : "removed")}");
                };
            });
        });
        _icon.MenuItems.AddSeparator();
        _icon.MenuItems.AddItem(x =>
        {
            x.Text = $"Version {TempoOutlookSyncContext.Version}";
            x.Clicked = _ => ShellOpen($"https://github.com/BlyZeDev/{nameof(TempoOutlookSync)}");
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
        _logger.Log += OnLog;
        _config.ConfigurationReload += OnConfigurationReload;
        
        try
        {
            //TODO
        }
        catch (OperationCanceledException) { }
    }

    public void Dispose()
    {
        _icon.Dispose();
        _cts.Dispose();
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

    private async void OnConfigurationReload(ConfigChangedEventArgs args)
    {
        if (!args.OldConfig.UserId.Equals(args.NewConfig.UserId, StringComparison.Ordinal)
            || !args.OldConfig.ApiToken.Equals(args.NewConfig.ApiToken, StringComparison.Ordinal)) await PerformManualSyncAsync();
    }

    private async Task SyncTempoToOutlookAsync()
    {
        if (Interlocked.Exchange(ref isSyncing, true)) return;

        try
        {
            await _tempo.ThrowIfCantConnect();

            _logger.LogInfo("Sync started");

            var today = DateTime.Today;
            var todayAddYear = today.AddYears(1);

            var existingTempoAppointments = _outlook.GetOutlookTempoAppointments(today);

            var changeCount = 0;
            await foreach (var entry in _tempo.GetPlannerEntriesAsync(today, todayAddYear))
            {
                var entryId = entry.Id.ToString();
                var needsCreation = true;

                if (existingTempoAppointments.TryGetValue(entryId, out var appointments))
                {
                    needsCreation = appointments.Any(item => _outlook.DeleteIfOutdated(item, entry.LastUpdated));
                    existingTempoAppointments.Remove(entryId);
                }

                if (!needsCreation) continue;

                changeCount++;
                switch (entry.RecurrenceRule)
                {
                    case RecurrenceRule.Never when entry.End.Date >= entry.Start.Date:
                        _outlook.SaveNonRecurring(entry);
                        break;

                    case RecurrenceRule.Weekly:
                    case RecurrenceRule.BiWeekly:
                        _outlook.SaveWeeklyRecurring(entry);
                        break;

                    case RecurrenceRule.Monthly when entry.End.Day != entry.Start.Day:
                        _outlook.SaveMonthlyRecurrence(entry);
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
                    obsoleteAppointment.Delete();
                }
            }

            _logger.LogInfo($"Synced {changeCount} item{(changeCount == 1 ? char.MinValue : 's')}, next sync in {Interval.TotalMinutes:F2} minutes");
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
            Interlocked.Exchange(ref isSyncing, false);
        }
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

    private static void ShellOpen(string fileName)
    {
        using (var process = new Process())
        {
            process.StartInfo = new ProcessStartInfo
            {
                UseShellExecute = true,
                FileName = fileName
            };
            process.Start();
        }
    }
}