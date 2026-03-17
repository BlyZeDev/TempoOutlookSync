namespace TempoOutlookSync.Services;

using CsToml;
using CsToml.Extensions.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Primitives;
using System.Text;
using TempoOutlookSync.Common;

public sealed class ConfigurationHandler : IDisposable
{
    private static readonly CsTomlSerializerOptions TomlOptions = CsTomlSerializerOptions.Default with
    {
        SerializeOptions = new SerializeOptions
        {
            ArrayStyle = TomlArrayStyle.Header
        }
    };
    private static UserSettings DefaultUserSettings = new UserSettings
    {
        Email = "",
        JiraApiToken = "",
        UserId = "",
        TempoApiToken = ""
    };
    private static CategorySettings DefaultCategories = new CategorySettings
    {
        Categories =
        [
            new Category
            {
                Name = "Intern",
                Color = OutlookColor.Purple,
                JQL = "issuekey = EDOCSE-136"
            },
            new Category
            {
                Name = "Support - Warte auf Kunde",
                Color = OutlookColor.DarkBlue,
                JQL = "(issuetype = Support or project = \"edoc Kundenportal\") and status = \"Waiting for Customer\""
            },
            new Category
            {
                Name = "Support - In Arbeit",
                Color = OutlookColor.Teal,
                JQL = "(issuetype = Support or project = \"edoc Kundenportal\") and status = \"In Progress\""
            },
            new Category
            {
                Name = "Support - Andere",
                Color = OutlookColor.Blue,
                JQL = "(issuetype = Support or project = \"edoc Kundenportal\") and statusCategory != Done"
            },
            new Category
            {
                Name = "Kundenprojekt - Aufgabe Kunde/Warte auf 3rd Level",
                Color = OutlookColor.Orange,
                JQL = "(category = BC or project = SP) and status = \"Aufgabe Kunde\" or status = \"Waiting for 3rd level\""
            },
            new Category
            {
                Name = "Kundenprojekt - In Arbeit/Aufgabe Edoc",
                Color = OutlookColor.DarkOrange,
                JQL = "(category = BC or project = SP) and status = \"In Progress\" or status = \"Aufgabe edoc\""
            },
            new Category
            {
                Name = "Kundenprojekt - Andere",
                Color = OutlookColor.Red,
                JQL = "(category = BC or project = SP) and statusCategory != Done"
            }
        ]
    };

    private readonly TempoOutlookSyncContext _context;
    private readonly ILogger _logger;

    private readonly IConfigurationRoot _root;
    private readonly IDisposable _reloadToken;

    public UserSettings UserSettings { get; private set; }

    public CategorySettings CategorySettings { get; private set; }

    public event Action<ObjectChangedEventArgs<UserSettings>>? UserSettingsChanged;
    public event Action<ObjectChangedEventArgs<CategorySettings>>? CategoriesChanged;

    public ConfigurationHandler(TempoOutlookSyncContext context, ILogger logger)
    {
        _context = context;
        _logger = logger;

        CreateDefaultUserSettings();

        _root = new ConfigurationBuilder()
            .SetBasePath(_context.AppFilesDirectory)
            .AddTomlFile(TempoOutlookSyncContext.UserSettingsFileName, true, true)
            .AddTomlFile(TempoOutlookSyncContext.CategoriesFileName, true, true)
            .SetFileLoadExceptionHandler(x => _logger.LogError(x.Exception.Message, x.Exception))
            .Build();

        UserSettings = LoadSettings();
        CategorySettings = LoadCategories();

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

    private UserSettings LoadSettings() => _root.Get<UserSettings>() ?? DefaultUserSettings;

    private CategorySettings LoadCategories() => _root.Get<CategorySettings>() ?? DefaultCategories;

    private void CreateDefaultUserSettings()
    {
        if (File.Exists(_context.UserSettingsPath)) return;

        using (var writer = new StreamWriter(_context.UserSettingsPath, false, Encoding.UTF8))
        {
            writer.WriteLine($"# {nameof(TempoOutlookSync)} user settings");
            writer.WriteLine();
            writer.WriteLine($"# Documentation can be found here: {_context.HelpUrl}");
            writer.WriteLine();

            writer.Write(CsTomlSerializer.Serialize(DefaultUserSettings, TomlOptions));
        }
    }

    private void CreateDefaultCategories()
    {
        if (File.Exists(_context.CategoriesPath)) return;

        using (var writer = new StreamWriter(_context.CategoriesPath, false, Encoding.UTF8))
        {
            writer.WriteLine($"# {nameof(TempoOutlookSync)} categories");
            writer.WriteLine();
            writer.WriteLine($"# Documentation can be found here: {_context.HelpUrl}");
            writer.WriteLine();

            writer.Write(CsTomlSerializer.Serialize(DefaultCategories, TomlOptions));
        }
    }

    public void Dispose() => _reloadToken.Dispose();
}