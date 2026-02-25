namespace TempoOutlookSync.Services;

using System.Diagnostics;
using System.Runtime;
using TempoOutlookSync.Common;

public sealed class MemoryMonitor : IDisposable
{
    private static readonly TimeSpan Interval = TimeSpan.FromSeconds(30);

    private readonly ILogger _logger;

    private CancellationTokenSource? cts;
    private Task? backgroundTask;

    public MemoryMonitor(ILogger logger) => _logger = logger;

    public async Task RunAsync(CancellationToken cancellationToken)
    {
        await StopAsync();

        if (backgroundTask is null)
        {
            cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            backgroundTask = MonitorLoopAsync(cts.Token);
        }
    }

    public void Dispose()
    {
        cts?.Cancel();
        cts?.Dispose();

        if (backgroundTask is not null)
        {
            while (!backgroundTask.IsCompleted) { }
            backgroundTask.Dispose();
        }
    }

    private async Task StopAsync()
    {
        if (cts is null || backgroundTask is null) return;

        await cts.CancelAsync();
        await backgroundTask;

        cts.Dispose();
        backgroundTask.Dispose();

        cts = null;
        backgroundTask = null;
    }

    private async Task MonitorLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            using (var process = Process.GetCurrentProcess())
            {
                using (var timer = new PeriodicTimer(Interval))
                {
                    while (await timer.WaitForNextTickAsync(cancellationToken))
                    {
                        LogMemoryMetrics(process);
                    }
                }
            }
        }
        catch (OperationCanceledException)
        {
            _logger.LogDebug("The memory monitoring ended");
        }
        catch (Exception ex)
        {
            _logger.LogDebug("The memory monitor crashed unexpectedly", ex);
        }
    }

    private void LogMemoryMetrics(Process process)
    {
        process.Refresh();

        var gcInfo = GC.GetGCMemoryInfo();
        
        //Just an example make it more useful
        _logger.LogDebug(
            $"""

            -- Memory --
            Managed: {Util.FormatBytes(GC.GetTotalMemory(false))}
            WorkingSet: {Util.FormatBytes(process.WorkingSet64)}
            Private: {Util.FormatBytes(process.PrivateMemorySize64)}

            Gen Collections:
                Gen0: {GC.CollectionCount(0)}
                Gen1: {GC.CollectionCount(1)}
                Gen2: {GC.CollectionCount(2)}

            Heap: {Util.FormatBytes(gcInfo.HeapSizeBytes)}
            Fragmented: {Util.FormatBytes(gcInfo.FragmentedBytes)}

            GC Mode: {GCSettings.LatencyMode}
            """);
    }
}