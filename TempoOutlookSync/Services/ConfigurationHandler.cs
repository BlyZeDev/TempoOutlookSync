namespace TempoOutlookSync.Services;

using CsToml;
using CsToml.Extensions.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Primitives;
using System.Text;
using TempoOutlookSync.Common;

public sealed class ConfigurationHandler : IDisposable
{
    private readonly ILogger _logger;

    private readonly IConfigurationRoot _root;
    private readonly IDisposable _reloadToken;

    public UserSettings UserSettings { get; private set; }

    public event Action<ObjectChangedEventArgs<UserSettings>>? UserSettingsChanged;

    public ConfigurationHandler(TempoOutlookSyncContext context, ILogger logger)
    {
        _logger = logger;

        CreateDefaultUserSettings(context.UserSettingsPath);

        _root = new ConfigurationBuilder()
            .SetBasePath(context.AppFilesDirectory)
            .AddTomlFile(TempoOutlookSyncContext.UserSettingsFileName, true, true)
            .SetFileLoadExceptionHandler(x => _logger.LogError(x.Exception.Message, x.Exception))
            .Build();

        UserSettings = LoadSettings();

        _reloadToken = ChangeToken.OnChange(_root.GetReloadToken, Reload);
    }

    private void Reload()
    {
        var oldSettings = UserSettings;
        UserSettings = LoadSettings();
        UserSettingsChanged?.Invoke(new ObjectChangedEventArgs<UserSettings>
        {
            Old = oldSettings,
            New = UserSettings
        });
        _logger.LogDebug("The configuration root was changed");
    }

    private UserSettings LoadSettings()
    {
        var settings = new UserSettings();
        _root.Bind(settings);
        return settings;
    }

    private static void CreateDefaultUserSettings(string userSettingsPath)
    {
        if (File.Exists(userSettingsPath)) return;

        using (var writer = new StreamWriter(userSettingsPath, false, Encoding.UTF8))
        {
            writer.WriteLine($"# {nameof(TempoOutlookSync)} user settings");
            writer.WriteLine();
            writer.WriteLine("# Please provide your Jira Email, found here: Click your profile picture at the top right in Jira");
            writer.WriteLine("# Please provide your Jira API Token, found here: https://docs.adaptavist.com/w4j/latest/quick-configuration-guide/add-sources/how-to-generate-jira-api-token");
            writer.WriteLine("# Please provide your Jira User Id, found here: https://www.storylane.io/tutorials/how-to-find-user-id-in-jira");
            writer.WriteLine("# Please provide your Tempo API Token, found here: https://help.tempo.io/timesheets/latest/using-rest-api-integrations");

            writer.Write(CsTomlSerializer.Serialize(new UserSettings()));
        }
    }

    public void Dispose() => _reloadToken.Dispose();
}