namespace TempoOutlookSync.Services;

using DotTray;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
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

    public ServiceRunner(ILogger logger, TempoOutlookSyncContext context, ConfigurationHandler config, TempoClient tempo, OutlookClient outlook)
    {
        _logger = logger;
        _context = context;
        _config = config;
        _tempo = tempo;
        _outlook = outlook;

        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnhandledTaskException;

        _cts = new CancellationTokenSource();
        _icon = NotifyIcon.Run(_context.IcoHandle, _cts.Token, x =>
        {
            x.BackgroundHoverColor = new TrayColor(187, 65, 203);
            x.BackgroundDisabledColor = new TrayColor(80, 80, 80);
            x.TextDisabledColor = new TrayColor(40, 40, 40);
        });
        _icon.SetToolTip($"{nameof(TempoOutlookSync)} - Version {TempoOutlookSyncContext.Version}");

        _icon.MenuItems.AddItem(x =>
        {
            x.Text = $"{nameof(TempoOutlookSync)} - Version {TempoOutlookSyncContext.Version}";
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
        try
        {
            using (var timer = new PeriodicTimer(Interval))
            {
                await PerformSync();

                while (await timer.WaitForNextTickAsync(_cts.Token))
                {
                    await PerformSync();
                }
            }
        }
        catch (OperationCanceledException) { }
    }

    public void Dispose()
    {
        _cts.Dispose();
    }

    private async Task PerformSync()
    {
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
        catch (Exception ex)
        {
            _logger.LogError("Sync failed", ex);
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