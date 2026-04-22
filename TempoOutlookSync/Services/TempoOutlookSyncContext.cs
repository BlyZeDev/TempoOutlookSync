namespace TempoOutlookSync.Services;

using System.Drawing;
using System.Reflection;
using System.Text;
using TempoOutlookSync.Common;

public sealed class TempoOutlookSyncContext : IDisposable
{
    public const string UserSettingsFileName = "usersettings.toml";
    public const string CategoriesFileName = "categories.toml";

    public const string ApplicationMutexId = $@"Global\{{{nameof(TempoOutlookSync)}-07863666-66fb-41ba-9cc1-83725487810d}}";

    private readonly HashSet<string> _tempPaths;

    /// <summary>
    /// The base directory of the application
    /// </summary>
    public string ApplicationDirectory => AppContext.BaseDirectory;

    /// <summary>
    /// The full path to the .exe of this application
    /// </summary>
    public string ExecutablePath => Environment.ProcessPath ?? throw new ApplicationException("The path of the executable could not be found");

    /// <summary>
    /// The base directory to store application files
    /// </summary>
    public string AppFilesDirectory
    {
        get
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), nameof(TempoOutlookSync));
            Directory.CreateDirectory(path);
            return path;
        }
    }

    /// <summary>
    /// The handle to the to the default application icon
    /// </summary>
    public string DefaultIcoPath { get; }

    /// <summary>
    /// The handle to the busy application icon
    /// </summary>
    public string BusyIcoPath { get; }

    /// <summary>
    /// The path to the user settings
    /// </summary>
    public string UserSettingsPath => Path.Combine(AppFilesDirectory, UserSettingsFileName);

    /// <summary>
    /// The path to the categories settings
    /// </summary>
    public string CategoriesPath => Path.Combine(AppFilesDirectory, CategoriesFileName);

    /// <summary>
    /// The base directory for all log files
    /// </summary>
    public string LogDirectory
    {
        get
        {
            var path = Path.Combine(AppFilesDirectory, "Logs");
            Directory.CreateDirectory(path);
            return path;
        }
    }

    /// <summary>
    /// The author of the application
    /// </summary>
    public string Author => "BlyZeDev";

    /// <summary>
    /// The base url to the GitHub repository
    /// </summary>
    public string GitHubRepoUrl => $"https://github.com/{Author}/{nameof(TempoOutlookSync)}";

    /// <summary>
    /// The url to the documentation
    /// </summary>
    public string HelpUrl => "https://edocag.atlassian.net/wiki/x/7wnyhw";

    /// <summary>
    /// The application argument use to indicate a headless sync and abort
    /// </summary>
    public string HeadlessArgument => "--sync";

    /// <summary>
    /// <see langword="true"/> if the application is running in headless mode (without UI), otherwise <see langword="false"/>
    /// </summary>
    public bool IsHeadless { get; }

    public TempoOutlookSyncContext()
    {
        _tempPaths = [];

        var args = Environment.GetCommandLineArgs();
        IsHeadless = args.Length == 2 && args[1].Equals(HeadlessArgument, StringComparison.Ordinal);

        var icoPath = CreateIco("icon.ico");
        if (!File.Exists(icoPath)) icoPath = CreateFallbackIco();
        if (!File.Exists(icoPath)) throw new ApplicationException("No icon could be created");

        var busyIco = CreateIco("icon_working.ico");
        if (!File.Exists(busyIco)) busyIco = icoPath;

        DefaultIcoPath = icoPath;
        BusyIcoPath = busyIco;
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

        var deletionBaseline = DateTime.UtcNow.AddDays(-7);
        foreach (var logFile in Directory.EnumerateFiles(LogDirectory, "*.log", SearchOption.TopDirectoryOnly))
        {
            try
            {
                if (File.GetLastWriteTimeUtc(logFile) < deletionBaseline)
                {
                    File.Delete(logFile);
                }
            }
            catch (Exception) { }
        }

        GC.SuppressFinalize(this);
    }

    private unsafe string? CreateIco(string resourceName)
    {
        using (var icoStream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"{nameof(TempoOutlookSync)}.{resourceName}"))
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