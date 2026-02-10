namespace TempoOutlookSync;

using TempoOutlookSync.Common;
using TempoOutlookSync.Services;

sealed class Program
{
    public static bool IsDebug
#if DEBUG
        => true;
#else
        => false;
#endif

    static void Main()
    {
        using (var provider = new ServiceProvider())
        {
            provider.GetService<ILogger>().LogInfo($"{nameof(TempoOutlookSync)} has started");

            using (var guard = provider.GetService<StartupGuard>())
            {
                if (!guard.WaitForAccess())
                {
                    provider.GetService<ILogger>().LogCritical($"{nameof(TempoOutlookSync)} is already running", null);
                    Environment.FailFast($"{nameof(TempoOutlookSync)} is already running");
                }

                provider.GetService<ServiceRunner>().RunAsync().GetAwaiter().GetResult();
            }
        }
    }
}