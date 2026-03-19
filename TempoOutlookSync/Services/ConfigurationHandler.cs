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
    private static readonly UserSettings DefaultUserSettings = new UserSettings
    {
        Email = "",
        JiraApiToken = "",
        UserId = "",
        TempoApiToken = ""
    };
    private static readonly CategorySettings DefaultCategories = new CategorySettings
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
                JQL = "(category = BC or project = SP or project = EDOCPMO) and status = \"Aufgabe Kunde\" or status = \"Waiting for 3rd level\""
            },
            new Category
            {
                Name = "Kundenprojekt - In Arbeit/Aufgabe Edoc",
                Color = OutlookColor.DarkOrange,
                JQL = "(category = BC or project = SP or project = EDOCPMO) and status = \"In Progress\" or status = \"Aufgabe edoc\""
            },
            new Category
            {
                Name = "Kundenprojekt - Andere",
                Color = OutlookColor.Red,
                JQL = "(category = BC or project = SP or project = EDOCPMO) and statusCategory != Done"
            }
        ]
    };

    private readonly TempoOutlookSyncContext _context;
    private readonly ILogger _logger;

    private readonly IConfigurationRoot _root;
    private readonly IDisposable _reloadToken;

    public UserSettings UserSettings { get; private set; }
    public CategorySettings CategorySettings { get; private set; }

    public ConfigurationHandler(TempoOutlookSyncContext context, ILogger logger)
    {
        _context = context;
        _logger = logger;

        TryCreateDefaultUserSettings();
        TryCreateDefaultCategories();

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
        UserSettings = LoadSettings();
        CategorySettings = LoadCategories();

        _logger.LogDebug("The configuration was changed");
    }

    private UserSettings LoadSettings()
    {
        if (!File.Exists(_context.UserSettingsPath)) return DefaultUserSettings;

        return _root.Get<UserSettings>() ?? DefaultUserSettings;
    }

    private CategorySettings LoadCategories()
    {
        if (!File.Exists(_context.CategoriesPath)) return DefaultCategories;

        return _root.Get<CategorySettings>() ?? DefaultCategories;
    }

    private void TryCreateDefaultUserSettings()
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

    private void TryCreateDefaultCategories()
    {
        if (File.Exists(_context.CategoriesPath)) return;

        using (var writer = new StreamWriter(_context.CategoriesPath, false, Encoding.UTF8))
        {
            writer.WriteLine($"# {nameof(TempoOutlookSync)} categories");
            writer.WriteLine("# Don't edit this file unless you know what you're doing!");
            writer.WriteLine();
            writer.WriteLine("# It is recommended to stop the application before changing this configuration");
            writer.WriteLine($"# Documentation can be found here: {_context.HelpUrl}");
            writer.WriteLine();

            writer.Write(CsTomlSerializer.Serialize(DefaultCategories, TomlOptions));
        }
    }

    public void Dispose() => _reloadToken.Dispose();
}