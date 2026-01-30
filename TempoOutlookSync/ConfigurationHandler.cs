using System;
using System.IO;
using System.Text.Json;

namespace TempoOutlookSync
{
    public sealed class ConfigurationHandler
    {
        private static readonly string ConfigPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "outlooksync.json");

        public Configuration Configuration { get; private set; }

        public bool IsValid => Configuration != null;

        private ConfigurationHandler(Configuration configuration)
        {
            Configuration = configuration;
        }

        public void SetConfiguration(Configuration configuration)
        {
            TrySetConfiguration(configuration);
            Configuration = TryGetConfiguration();
        }

        public void DeleteConfiguration()
        {
            try
            {
                File.Delete(ConfigPath);
                TryGetConfiguration();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public static ConfigurationHandler Initialize() => new ConfigurationHandler(TryGetConfiguration());

        private static Configuration TryGetConfiguration()
        {
            try
            {
                using (var fileStream = new FileStream(ConfigPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return JsonSerializer.Deserialize<Configuration>(fileStream);
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TrySetConfiguration(Configuration config)
        {
            try
            {
                using (var fileStream = new FileStream(ConfigPath, FileMode.Create, FileAccess.Write, FileShare.Read))
                {
                    JsonSerializer.Serialize(fileStream, config);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
