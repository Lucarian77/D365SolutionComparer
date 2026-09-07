using System;
using System.IO;
using System.Xml.Serialization;

namespace D365SolutionComparer
{
    [Serializable]
    public sealed class Settings
    {
        private static readonly string SettingsFolderPath =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AdrianLucaci",
                "D365SolutionComparer");

        private static readonly string SettingsFilePath =
            Path.Combine(SettingsFolderPath, "settings.xml");

        public bool FilterAll { get; set; } = true;
        public bool FilterMatch { get; set; }
        public bool FilterVersionMismatch { get; set; }
        public bool FilterPublisherMismatch { get; set; }
        public bool FilterDisplayNameMismatch { get; set; }
        public bool FilterPackageTypeDifference { get; set; }
        public bool FilterMultipleDifferences { get; set; }
        public bool FilterMissingInSource { get; set; }
        public bool FilterMissingInTarget { get; set; }

        public bool ShowManagedUnmanagedOnly { get; set; }
        public bool ShowChangedOnly { get; set; }

        public static Settings Load()
        {
            try
            {
                if (!File.Exists(SettingsFilePath))
                {
                    return new Settings();
                }

                var serializer = new XmlSerializer(typeof(Settings));

                using (var stream = File.OpenRead(SettingsFilePath))
                {
                    var loaded = serializer.Deserialize(stream) as Settings;
                    return loaded ?? new Settings();
                }
            }
            catch
            {
                return new Settings();
            }
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(SettingsFolderPath);

                var serializer = new XmlSerializer(typeof(Settings));

                using (var stream = File.Create(SettingsFilePath))
                {
                    serializer.Serialize(stream, this);
                }
            }
            catch
            {
                // Ignore save failures. Settings persistence should never block the tool.
            }
        }
    }
}