namespace TempoOutlookSync;

using TempoOutlookSync.Services;
using Velopack;

sealed class Program
{
    static void Main()
    {
        VelopackApp.Build().Run();

        using (var provider = new ServiceProvider())
        {
            if (provider.GetService<TempoOutlookSyncContext>().IsHeadless)
            {
                provider.GetService<SynchronizationHandler>().ExecuteAsync().GetAwaiter().GetResult();
                return;
            }

            provider.GetService<SynchronizationScheduler>().Register();

            using (var guard = provider.GetService<StartupGuard>())
            {
                if (!guard.WaitForAccess())
                {
                    provider.GetService<ILogger>().LogCritical($"{nameof(TempoOutlookSync)} is already running", null);
                    Environment.FailFast($"{nameof(TempoOutlookSync)} is already running");
                }

                provider.GetService<ILogger>().LogInfo($"{nameof(TempoOutlookSync)} {provider.GetService<UpdateHandler>().VersionString} is now running");

                provider.GetService<UpdateHandler>().UpdateAndRestartIfAvailable();
                provider.GetService<TrayService>().RunAsync().GetAwaiter().GetResult();
                provider.GetService<UpdateHandler>().UpdateAndExitIfAvailable();
            }
        }
    }
}