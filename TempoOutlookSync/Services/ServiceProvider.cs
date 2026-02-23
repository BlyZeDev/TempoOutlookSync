namespace TempoOutlookSync.Services;

using Jab;

[ServiceProvider]
[Singleton<StartupGuard>]
[Singleton<ServiceRunner>]
[Singleton<ILoggerTarget, FileLogger>]
[Singleton<ILogger, LoggerForwarder>]
[Singleton<TempoOutlookSyncContext>]
[Singleton<AppConfiguration>]
[Singleton<TempoApiClient>]
[Singleton<JiraApiClient>]
[Singleton<OutlookClient>]
public sealed partial class ServiceProvider;