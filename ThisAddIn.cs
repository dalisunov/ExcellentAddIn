using System;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcellentAddIn.Core;
using ExcellentAddIn.Interfaces;
using ExcellentAddIn.Configuration;
using ExcellentAddIn.Events;
using Microsoft.Office.Core;

namespace ExcellentAddIn
{
    public partial class ThisAddIn
    {
        private CustomRibbon ribbon;
        private ILogger _logger;
        private IErrorHandler _errorHandler;

        // Добавляем конфигурацию
        private ConfigurationManager _configManager;
        private AppSettings _appSettings;
        private UserSettings _userSettings;

        // Добавляем EventManager
        private EventManager _eventManager;

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new CustomRibbon();
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Инициализация логгера и обработчика ошибок
            _logger = new Logger();
            _errorHandler = new ErrorManager(_logger);

            // Инициализация ConfigurationManager
            string appSettingsPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcellentAddIn",
                "AppSettings.xml");

            string userSettingsPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcellentAddIn",
                "UserSettings.xml");

            _configManager = new ConfigurationManager(appSettingsPath, userSettingsPath);

            // Загружаем настройки
            _appSettings = _configManager.LoadAppSettings();
            _userSettings = _configManager.LoadUserSettings();

            // Проверяем настройки и при необходимости заполняем дефолтными значениями
            if (string.IsNullOrEmpty(_appSettings.Version))
            {
                _appSettings.Version = "1.0";
                _configManager.SaveAppSettings(_appSettings);
            }

            if (string.IsNullOrEmpty(_userSettings.UserName))
            {
                _userSettings.UserName = Environment.UserName;
                _configManager.SaveUserSettings(_userSettings);
            }

            _logger.LogInfo($"Startup: Application version {_appSettings.Version}, User: {_userSettings.UserName}.");

            // Подключаем менеджер событий
            _eventManager = new EventManager(this.Application);
            _eventManager.StartListening();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Выгружаем Add-In
            _logger.LogInfo("ThisAddIn_Shutdown: выгрузка Add-In.");

            // Сохраняем актуальные настройки перед выходом (если что-то менялось)
            _configManager.SaveAppSettings(_appSettings);
            _configManager.SaveUserSettings(_userSettings);

            // Отписываемся от событий
            if (_eventManager != null)
            {
                _eventManager.StopListening();
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
