namespace TempoOutlookSync.Services;

using DotTray;
using System;
using System.Diagnostics.CodeAnalysis;
using TempoOutlookSync.Common;

public sealed class TrayService : IDisposable
{
    private readonly ILogger _logger;
    private readonly TempoOutlookSyncContext _context;
    private readonly UpdateHandler _updater;
    private readonly ConfigurationHandler _config;
    private readonly SynchronizationScheduler _scheduler;
    private readonly OutlookComClient _outlook;

    private readonly CancellationTokenSource _cts;
    private readonly NotifyIcon _icon;
    private readonly MenuItem _syncNowMenuItem;
    private readonly MenuItem _nextSyncMenuItem;
    private readonly MenuItem _debugMenuItem;

    public TrayService(ILogger logger, TempoOutlookSyncContext context, UpdateHandler updater, ConfigurationHandler config, SynchronizationScheduler scheduler, OutlookComClient outlook)
    {
        _logger = logger;
        _context = context;
        _updater = updater;
        _config = config;
        _scheduler = scheduler;
        _outlook = outlook;

        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnhandledTaskException;

        _cts = new CancellationTokenSource();
        _icon = NotifyIcon.Run(_context.DefaultIcoPath, _cts.Token, x =>
        {
            x.BackgroundHoverColor = new TrayColor(218, 83, 225);
            x.BackgroundDisabledColor = new TrayColor(40, 40, 40);
            x.TextDisabledColor = new TrayColor(180, 180, 180);
        }, x => x.LineThickness = 1.2f);
        _icon.SetFontSize(16f);
        _icon.SetToolTip($"{nameof(TempoOutlookSync)} - Version {_updater.VersionString}");

        _syncNowMenuItem = _icon.MenuItems.AddItem(x =>
        {
            x.Text = "Sync now";
            x.Clicked = _ => _scheduler.Run();
        });
        _nextSyncMenuItem = _icon.MenuItems.AddItem(x =>
        {
            x.Text = $"Next Sync in {Util.FormatTime(GetRemainingToNextSync())}";
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
                x.Clicked = _ => Util.ShellOpen(_context.HelpUrl);
            });
        });
        _debugMenuItem = _icon.MenuItems.AddItem(x =>
        {
            x.Text = "Debug";
            x.BackgroundHoverColor = new TrayColor(255, 0, 0);
            x.SubMenu.AddItem(x =>
            {
                x.Text = "Delete ALL synced items";
                x.BackgroundHoverColor = new TrayColor(255, 0, 0);
                x.Clicked = async _ =>
                {
                    _scheduler.Disable();

                    try
                    {
                        _logger.LogInfo("Started deleting all synced entries");

                        await Task.Run(() =>
                        {
                            foreach (var appointment in _outlook.GetTempoAppointments())
                            {
                                _outlook.DeleteByEntryId(appointment.EntryId);
                            }

                            _outlook.PurgeTrashedTempoAppointments();
                        });

                        _logger.LogInfo("Finished deleting all synced entries");
                    }
                    finally
                    {
                        _scheduler.Enable();
                    }
                };
            });
        });
        _icon.MenuItems.AddSeparator();
        _icon.MenuItems.AddItem(x =>
        {
            x.Text = $"Version - {_updater.VersionString}";
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
        _icon.PopupShowing += OnPopupShowing;

        _icon.ShowBalloon(new BalloonNotification
        {
            Icon = BalloonNotificationIcon.User,
            Title = $"{nameof(TempoOutlookSync)}",
            Message = "is now running in the background"
        });

        try
        {
            await Task.Delay(Timeout.Infinite, _cts.Token);
        }
        catch (OperationCanceledException) { }

        _icon.PopupShowing -= OnPopupShowing;
        _logger.Log -= OnLog;
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

    private void OnSyncStateChange()
    {
        var info = _scheduler.GetInfo();
        _logger.LogDebug($"State changed to {info.State}");

        var isBusy = info.State is Microsoft.Win32.TaskScheduler.TaskState.Running;
        _syncNowMenuItem.IsDisabled = _debugMenuItem.IsDisabled = isBusy;

        OnPopupShowing(MouseButton.None);

        if (isBusy) _icon.SetIcon(_context.BusyIcoPath);
        else _icon.SetIcon(_context.DefaultIcoPath);
    }

    private void OnPopupShowing(MouseButton mouseButton)
    {
        var info = _scheduler.GetInfo();

        if (info.State is Microsoft.Win32.TaskScheduler.TaskState.Running) _nextSyncMenuItem.Text = "Syncing...";
        else _nextSyncMenuItem.Text = $"Next Sync in {Util.FormatTime(GetRemainingToNextSync())}";
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
        
        _context.WriteCrashLog(exception);
        Environment.FailFast(exception.Message, exception);
    }

    private TimeSpan GetRemainingToNextSync()
    {
        var remaining = _scheduler.GetInfo().NextRunTime - DateTime.UtcNow;
        return remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero;
    }
}