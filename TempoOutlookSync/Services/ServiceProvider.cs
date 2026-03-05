namespace TempoOutlookSync.Services;

using Jab;

[ServiceProvider]
[Singleton<StartupGuard>]
[Singleton<ServiceRunner>]
[Singleton<ILoggerTarget, FileLogger>]
[Singleton<ILogger, LoggerForwarder>]
[Singleton<TempoOutlookSyncContext>]
[Singleton<UpdateHandler>]
[Singleton<ConfigurationHandler>]
[Singleton<TempoApiClient>]
[Singleton<JiraApiClient>]
[Singleton<OutlookComClient>]
public sealed partial class ServiceProvider;