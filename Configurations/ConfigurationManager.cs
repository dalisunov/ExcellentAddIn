using System;
using System.IO;
using System.Xml.Serialization;

namespace ExcellentAddIn.Configuration
{
    /// <summary>
    /// Менеджер для загрузки и сохранения настроек (AppSettings, UserSettings)
    /// </summary>
    public class ConfigurationManager
    {
        private readonly string _appSettingsFilePath;
        private readonly string _userSettingsFilePath;

        public ConfigurationManager(string appSettingsFilePath, string userSettingsFilePath)
        {
            _appSettingsFilePath = appSettingsFilePath;
            _userSettingsFilePath = userSettingsFilePath;
        }

        public AppSettings LoadAppSettings()
        {
            return LoadSettings<AppSettings>(_appSettingsFilePath);
        }

        public UserSettings LoadUserSettings()
        {
            return LoadSettings<UserSettings>(_userSettingsFilePath);
        }

        public void SaveAppSettings(AppSettings settings)
        {
            SaveSettings(settings, _appSettingsFilePath);
        }

        public void SaveUserSettings(UserSettings settings)
        {
            SaveSettings(settings, _userSettingsFilePath);
        }

        private T LoadSettings<T>(string path) where T : class, new()
        {
            if (!File.Exists(path))
                return new T();

            try
            {
                using (var stream = File.OpenRead(path))
                {
                    var serializer = new XmlSerializer(typeof(T));
                    return serializer.Deserialize(stream) as T;
                }
            }
            catch (Exception)
            {
                // При ошибке десериализации возвращаем объект по умолчанию, или логируем ошибку
                return new T();
            }
        }

        private void SaveSettings<T>(T settings, string path) where T : class
        {
            try
            {
                using (var stream = File.Create(path))
                {
                    var serializer = new XmlSerializer(typeof(T));
                    serializer.Serialize(stream, settings);
                }
            }
            catch (Exception ex)
            {
                // Логируем ошибку или обрабатываем
                Console.WriteLine($"Error saving settings: {ex.Message}");
            }
        }
    }
}
