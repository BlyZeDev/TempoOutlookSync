namespace TempoOutlookSync;

using TempoOutlookSync.Services;
using Velopack;

sealed class Program
{
    static void Main()
    {
        using (var provider = new ServiceProvider())
        {
            using (var guard = provider.GetService<StartupGuard>())
            {
                if (!guard.WaitForAccess())
                {
                    provider.GetService<ILogger>().LogCritical($"{nameof(TempoOutlookSync)} is already running", null);
                    Environment.FailFast($"{nameof(TempoOutlookSync)} is already running");
                }

                VelopackApp.Build().Run();

                provider.GetService<ILogger>().LogInfo($"{nameof(TempoOutlookSync)} {provider.GetService<UpdateHandler>().Version} is now running");

                provider.GetService<UpdateHandler>().UpdateAndRestartIfAvailable();
                provider.GetService<ServiceRunner>().RunAsync().GetAwaiter().GetResult();
                provider.GetService<UpdateHandler>().UpdateAndExitIfAvailable();
            }
        }
    }
}