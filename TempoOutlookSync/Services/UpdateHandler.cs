namespace TempoOutlookSync.Services;

using Velopack;
using Velopack.Sources;

public sealed class UpdateHandler
{
    private readonly UpdateManager _update;

    public string Version => _update.CurrentVersion is null ? "Invalid" : _update.CurrentVersion.ToFullString();

    public UpdateHandler(TempoOutlookSyncContext context)
    {
        _update = new UpdateManager(new GithubSource(context.GitHubRepoUrl, null, false), new UpdateOptions
        {
            AllowVersionDowngrade = false
        });
    }

    public void UpdateAndRestartIfAvailable()
    {
        if (!_update.IsInstalled) return;

        var available = _update.CheckForUpdates();
        if (available is null || available.IsDowngrade) return;

        _update.DownloadUpdates(available);
        _update.ApplyUpdatesAndRestart(available);
    }

    public void UpdateAndExitIfAvailable()
    {
        if (!_update.IsInstalled) return;

        var available = _update.CheckForUpdates();
        if (available is null || available.IsDowngrade) return;

        _update.DownloadUpdates(available);
        _update.ApplyUpdatesAndExit(available);
    }
}