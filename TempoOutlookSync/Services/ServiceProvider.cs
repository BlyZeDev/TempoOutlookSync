namespace TempoOutlookSync.Services;

using Jab;

[ServiceProvider]
[Singleton<StartupGuard>]
[Singleton<ServiceRunner>]
[Singleton<ILoggerTarget, ConsoleLogger>]
[Singleton<ILoggerTarget, FileLogger>]
[Singleton<ILogger, LoggerForwarder>]
[Singleton<TempoOutlookSyncContext>]
[Singleton<ConfigurationHandler>]
[Singleton<TempoApiClient>]
[Singleton<JiraApiClient>]
[Singleton<OutlookClient>]
public sealed partial class ServiceProvider;