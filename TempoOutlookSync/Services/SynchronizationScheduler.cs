namespace TempoOutlookSync.Services;

using Microsoft.Win32.TaskScheduler;
using System.Diagnostics;
using TempoOutlookSync.Common;

public sealed class SynchronizationScheduler
{
    private const string TaskName = $"{nameof(TempoOutlookSync)}-{nameof(SynchronizationScheduler)}";
    private static readonly TimeSpan Interval = TimeSpan.FromMinutes(15);

    private readonly ILogger _logger;
    private readonly TempoOutlookSyncContext _context;
    private readonly UpdateHandler _update;

    public SynchronizationScheduler(ILogger logger, TempoOutlookSyncContext context, UpdateHandler update)
    {
        _logger = logger;
        _context = context;
        _update = update;
    }

    public void Register()
    {
        _logger.LogDebug($"Registering scheduler to run every {Util.FormatTime(Interval)}");

        var task = TaskService.Instance.NewTask();

        task.Settings.AllowDemandStart = true;
        task.Settings.AllowHardTerminate = true;
        task.Settings.DisallowStartIfOnBatteries = false;
        task.Settings.DisallowStartOnRemoteAppSession = false;
        task.Settings.MultipleInstances = TaskInstancesPolicy.IgnoreNew;
        task.Settings.Priority = ProcessPriorityClass.Normal;
        task.Settings.RestartCount = 0;
        task.Settings.RunOnlyIfIdle = false;
        task.Settings.RunOnlyIfNetworkAvailable = true;
        task.Settings.StartWhenAvailable = true;
        task.Settings.StopIfGoingOnBatteries = false;
        task.Settings.Volatile = false;
        task.Settings.WakeToRun = false;

        task.RegistrationInfo.Version = _update.Version ?? new Version();
        task.RegistrationInfo.Author = _context.Author;
        task.RegistrationInfo.Description = "This task is responsible for scheduling the synchronization from Tempo Capacity Planner to Outlook Calendar";
        task.RegistrationInfo.Date = DateTime.Now;
        task.RegistrationInfo.URI = _context.HelpUrl;
        task.RegistrationInfo.Source = _context.ExecutablePath;

        task.Triggers.Add(new TimeTrigger
        {
            Enabled = true,
            Repetition = new RepetitionPattern(Interval, TimeSpan.Zero),
            ExecutionTimeLimit = Interval * 0.9
        });

        task.Actions.Add(new ExecAction
        {
            Path = _context.ExecutablePath,
            WorkingDirectory = _context.ApplicationDirectory,
            Arguments = _context.HeadlessArgument
        });

        task.Principal.LogonType = TaskLogonType.InteractiveToken;
        task.Principal.RunLevel = TaskRunLevel.LUA;

        var registered = TaskService.Instance.RootFolder.RegisterTaskDefinition(TaskName, task);

        _logger.LogDebug($"Scheduler is registered here: {registered.Path}");
    }

    public void Delete()
    {
        TaskService.Instance.RootFolder.DeleteTask(TaskName, false);
        _logger.LogDebug($"Deleted to following scheduler: {TaskName}");
    }

    public void Run() => TaskService.Instance.GetTask(TaskName)?.Run();

    public void Enable()
    {
        var task = TaskService.Instance.GetTask(TaskName);
        if (task is null) return;

        task.Definition.Triggers.Add(new TimeTrigger
        {
            Enabled = true,
            Repetition = new RepetitionPattern(Interval, TimeSpan.Zero),
            ExecutionTimeLimit = Interval * 0.9
        });
        task.RegisterChanges();

        _logger.LogDebug($"Time Trigger added - The scheduler will run with the following interval, {Util.FormatTime(Interval)}");
    }

    public void Disable()
    {
        var task = TaskService.Instance.GetTask(TaskName);
        if (task is null) return;

        task.Definition.Triggers.Clear();
        task.RegisterChanges();

        _logger.LogDebug("All triggers removed - The scheduler won't run until activated again");
    }

    public ScheduledTaskInfo GetInfo()
    {
        var task = TaskService.Instance.GetTask(TaskName);

        return new ScheduledTaskInfo
        {
            IsActive = task.IsActive,
            State = task.State,
            NextRunTime = task.NextRunTime
        };
    }
}