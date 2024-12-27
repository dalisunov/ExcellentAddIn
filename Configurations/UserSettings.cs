namespace ExcellentAddIn.Configuration
{
    /// <summary>
    /// Настройки, специфичные для конкретного пользователя
    /// </summary>
    public class UserSettings
    {
        public string UserName { get; set; }
        public string PreferredTheme { get; set; }
        public bool RememberLastWorkbook { get; set; }

        // Другие настройки, которые могут меняться у разных пользователей
    }
}
