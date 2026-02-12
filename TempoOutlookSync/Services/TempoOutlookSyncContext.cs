namespace TempoOutlookSync.Services;

using System.Drawing;
using System.Reflection;
using System.Text;
using TempoOutlookSync.Common;

public sealed class TempoOutlookSyncContext : IDisposable
{
    public const string Version = "2.0.0";
    public const string ConfigFileName = "usersettings.json";

    private readonly HashSet<string> _tempPaths;

    /// <summary>
    /// The base directory of the application
    /// </summary>
    public string ApplicationDirectory { get; }

    /// <summary>
    /// The full path to the .exe of this application
    /// </summary>
    public string ExecutablePath { get; }

    /// <summary>
    /// The base directory to store application files
    /// </summary>
    public string AppFilesDirectory { get; }

    /// <summary>
    /// The handle to the to application icon
    /// </summary>
    public string IcoPath { get; }

    /// <summary>
    /// The path to the configuration
    /// </summary>
    public string ConfigurationPath { get; }

    /// <summary>
    /// The base directory for all log files
    /// </summary>
    public string LogDirectory { get; }

    public TempoOutlookSyncContext()
    {
        _tempPaths = [];

        ApplicationDirectory = AppContext.BaseDirectory;

        ExecutablePath = Environment.ProcessPath ?? throw new ApplicationException("The path of the executable could not be found");

        AppFilesDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), nameof(TempoOutlookSync));
        Directory.CreateDirectory(AppFilesDirectory);

        var icoPath = CreateMainIco();
        if (!File.Exists(icoPath)) icoPath = CreateFallbackIco();
        if (!File.Exists(icoPath)) throw new ApplicationException("No icon could be created");

        IcoPath = icoPath;

        ConfigurationPath = Path.Combine(AppFilesDirectory, ConfigFileName);

        LogDirectory = Path.Combine(AppFilesDirectory, "Logs");
        Directory.CreateDirectory(LogDirectory);
    }

    public string GetTempPath(string fileExtension)
    {
        string tempPath;
        do
        {
            tempPath = Path.ChangeExtension(Path.Combine(nameof(TempoOutlookSync), Path.GetTempPath(), Guid.CreateVersion7().ToString("N")), fileExtension);
        } while (!_tempPaths.Add(tempPath));

        return tempPath;
    }

    public string WriteCrashLog(Exception exception)
    {
        var crashLogPath = Path.Combine(LogDirectory, $"{nameof(TempoOutlookSync)}-Crash-{Util.GetFileNameTimestamp()}.log");

        var options = new FileStreamOptions
        {
            Access = FileAccess.Write,
            Mode = FileMode.Create,
            Options = FileOptions.WriteThrough,
            Share = FileShare.None
        };
        using (var writer = new StreamWriter(crashLogPath, Encoding.UTF8, options))
        {
            writer.Write(exception.ToString());
        }

        return crashLogPath;
    }

    public void Dispose()
    {
        foreach (var tempPath in _tempPaths)
        {
            if (File.Exists(tempPath))
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch (Exception) { }
            }
        }

        var deletetionBaseline = DateTime.UtcNow.AddDays(-7);
        foreach (var logFile in Directory.EnumerateFiles(LogDirectory, "*.log", SearchOption.TopDirectoryOnly))
        {
            if (File.GetCreationTimeUtc(logFile) < deletetionBaseline)
            {
                try
                {
                    File.Delete(logFile);
                }
                catch (Exception) { }
            }
        }

        GC.SuppressFinalize(this);
    }

    private unsafe string? CreateMainIco()
    {
        using (var icoStream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"{nameof(TempoOutlookSync)}.icon.ico"))
        {
            if (icoStream is null) return null;

            var tempPath = GetTempPath(".ico");

            using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                icoStream.CopyTo(fileStream);
                fileStream.Flush();
            }

            return tempPath;
        }
    }

    private string? CreateFallbackIco()
    {
        var tempPath = GetTempPath(".ico");

        using (var icon = SystemIcons.GetStockIcon(StockIconId.Error, StockIconOptions.SmallIcon))
        {
            using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                icon.Save(fileStream);
            }
        }

        return tempPath;
    }
}