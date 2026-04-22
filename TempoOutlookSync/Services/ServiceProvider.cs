namespace TempoOutlookSync.Services;

using Jab;

[ServiceProvider]
[Singleton<StartupGuard>]
[Singleton<SynchronizationScheduler>]
[Singleton<SynchronizationHandler>]
[Singleton<TrayService>]
[Singleton<ILoggerTarget, FileLogger>]
#if DEBUG
[Singleton<ILoggerTarget, DebugLogger>]
#endif
[Singleton<ILogger, LoggerForwarder>]
[Singleton<TempoOutlookSyncContext>]
[Singleton<UpdateHandler>]
[Singleton<ConfigurationHandler>]
[Singleton<TempoApiClient>]
[Singleton<JiraApiClient>]
[Singleton<OutlookComClient>]
public sealed partial class ServiceProvider;