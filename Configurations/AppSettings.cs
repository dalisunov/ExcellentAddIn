namespace ExcellentAddIn.Configuration
{
    /// <summary>
    /// Глобальные настройки, которые применяются ко всей надстройке
    /// </summary>
    public class AppSettings
    {
        /// <summary>
        /// Версия приложения (может храниться в XML, JSON и т.д.)
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// Включить/выключить логирование
        /// </summary>
        public bool EnableLogging { get; set; }

        /// <summary>
        /// Путь к файлу лога
        /// </summary>
        public string LogFilePath { get; set; }

        // Можно добавлять и другие настройки
        // (API ключи, URL сервисов, таймауты, и т.д.)

        /// <summary>
        /// Конструктор по умолчанию (можно задать дефолтные значения)
        /// </summary>
        public AppSettings()
        {
            Version = "1.0.0";
            EnableLogging = true;
            LogFilePath = Constants.DefaultLogFileName;
        }
    }
}
