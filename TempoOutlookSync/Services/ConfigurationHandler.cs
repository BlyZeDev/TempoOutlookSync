namespace TempoOutlookSync.Services;

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.Json;
using TempoOutlookSync.Common;

public sealed class ConfigurationHandler : IDisposable
{
    private static readonly Configuration _defaultConfig = new Configuration
    {
        UserId = "",
        ApiToken = ""
    };

    private readonly TempoOutlookSyncContext _context;
    private readonly ILogger _logger;
    private readonly FileSystemWatcher _watcher;

    public Configuration Current { get; private set; }

    public event Action<ConfigChangedEventArgs>? ConfigurationReload;

    public ConfigurationHandler(ILogger logger, TempoOutlookSyncContext context)
    {
        _logger = logger;
        _context = context;

        _watcher = new FileSystemWatcher
        {
            Path = _context.AppFilesDirectory,
            Filter = TempoOutlookSyncContext.ConfigFileName,
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.CreationTime | NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents = true
        };
        _watcher.Changed += OnConfigFileChange;
        _watcher.Created += OnConfigFileChange;
        _watcher.Deleted += OnConfigFileChange;
        _watcher.Renamed += OnConfigFileRenamed;
        _watcher.Error += OnWatcherError;

        Current = _defaultConfig;

        if (!File.Exists(_context.ConfigurationPath)) Write(_defaultConfig);
        Reload();
    }

    public void Write(Configuration config)
    {
        try
        {
            using (var writer = new StreamWriter(_context.ConfigurationPath, false, Encoding.UTF8))
            {
                var json = JsonSerializer.Serialize(config, ConfigurationJsonContext.Default.Configuration);

                writer.Write(json);
                writer.Flush();
            }

            _logger.LogDebug("Configuration was overwritten");
        }
        catch (IOException ex) when (IsFileLocked(ex))
        {
            _logger.LogDebug("The file is currently locked, ignoring.");
        }
    }

    public void Dispose()
    {
        _watcher.Changed -= OnConfigFileChange;
        _watcher.Created -= OnConfigFileChange;
        _watcher.Deleted -= OnConfigFileChange;
        _watcher.Renamed -= OnConfigFileRenamed;
        _watcher.Error -= OnWatcherError;
        _watcher.Dispose();

        GC.SuppressFinalize(this);
    }

    private void Reload()
    {
        try
        {
            var oldConfig = Current;

            using (var reader = new StreamReader(_context.ConfigurationPath, Encoding.UTF8))
            {
                var json = reader.ReadToEnd();
                Current = JsonSerializer.Deserialize(json, ConfigurationJsonContext.Default.Configuration) ?? throw new JsonException("The configuration can't be null");
            }

            ConfigurationReload?.Invoke(new ConfigChangedEventArgs
            {
                OldConfig = oldConfig,
                NewConfig = Current
            });

            _logger.LogInfo("Reloaded the configuration");
            _logger.LogDebug(Current.ToString() ?? "The configuration is <NULL>");
        }
        catch (IOException ex) when (IsFileLocked(ex))
        {
            _logger.LogDebug("The file is currently locked, ignoring.");
        }
        catch (JsonException ex)
        {
            _logger.LogWarning($"The configuration contains an error on line {(ex.LineNumber ?? -2) + 1}. Resetting to the last correct configuration");
            Write(Current);
        }
    }

    private void OnConfigFileChange(object sender, FileSystemEventArgs e)
    {
        switch (e.ChangeType)
        {
            case WatcherChangeTypes.Created or WatcherChangeTypes.Changed:
                _logger.LogDebug($"Configuration - {e.ChangeType}");
                Reload();
                break;

            case WatcherChangeTypes.Deleted:
                _logger.LogDebug($"Configuration - {e.ChangeType}");
                Write(_defaultConfig);
                break;
        }
    }

    private void OnConfigFileRenamed(object sender, RenamedEventArgs e)
    {
        _logger.LogDebug($"Configuration - {e.ChangeType}");

        if (_context.ConfigurationPath == e.FullPath) Reload();
        else Write(_defaultConfig);
    }

    private void OnWatcherError(object sender, ErrorEventArgs e) => _logger.LogError("The configuration file can't be monitored anymore. Any changes are not refreshed at runtime", e.GetException());

    private static bool IsFileLocked(IOException ex)
    {
        const int ERROR_SHARING_VIOLATION = 32;
        const int ERROR_LOCK_VIOLATION = 33;

        return (ex.HResult & 0xFFFF) is ERROR_SHARING_VIOLATION or ERROR_LOCK_VIOLATION;
    }
}