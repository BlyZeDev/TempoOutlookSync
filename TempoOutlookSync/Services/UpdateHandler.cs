namespace TempoOutlookSync.Services;

using Velopack;
using Velopack.Sources;

public sealed class UpdateHandler
{
    private readonly ILogger _logger;
    private readonly UpdateManager _manager;

    public Version? Version
    {
        get
        {
            var curVer = _manager.CurrentVersion;
            if (curVer is null) return null;

            return new Version(curVer.Major, curVer.Minor, curVer.Patch);
        }
    }
    public string? VersionString => _manager.CurrentVersion?.ToFullString();

    public UpdateHandler(ILogger logger, TempoOutlookSyncContext context)
    {
        _logger = logger;
        _manager = new UpdateManager(new GithubSource(context.GitHubRepoUrl, null, false), new UpdateOptions
        {
            AllowVersionDowngrade = false
        });
    }

    public void UpdateAndRestartIfAvailable()
    {
        try
        {
            _logger.LogDebug("Looking for updates");

            if (!_manager.IsInstalled) return;

            var available = _manager.CheckForUpdates();
            if (available is null || available.IsDowngrade) return;

            _manager.DownloadUpdates(available);
            _manager.ApplyUpdatesAndRestart(available);
        }
        catch (Exception ex)
        {
            _logger.LogError("Something went wrong trying to update", ex);
        }
    }

    public void UpdateAndExitIfAvailable()
    {
        try
        {
            _logger.LogDebug("Looking for updates");

            if (!_manager.IsInstalled) return;

            var available = _manager.CheckForUpdates();
            if (available is null || available.IsDowngrade) return;

            _manager.DownloadUpdates(available);
            _manager.ApplyUpdatesAndExit(available);
        }
        catch (Exception ex)
        {
            _logger.LogError("Something went wrong trying to update", ex);
        }
    }
}