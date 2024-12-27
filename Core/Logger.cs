using System;
using System.Diagnostics;
using System.IO;
using ExcellentAddIn.Interfaces;
using ExcellentAddIn.Configuration;

namespace ExcellentAddIn.Core
{
    /// <summary>
    /// Реализует функционал логирования с поддержкой вывода в Debug и файл.
    /// Потокобезопасная реализация для сценариев параллельного логирования.
    /// </summary>
    public class Logger : ILogger
    {
        private readonly string _logFilePath;
        private readonly object _lockObject = new object();
        private readonly bool _enableFileLogging;

        public Logger(string logFilePath = null)
        {
            _logFilePath = logFilePath ?? Constants.DefaultLogFileName;
            _enableFileLogging = !string.IsNullOrWhiteSpace(_logFilePath);

            // Проверяем существование директории для логов
            if (_enableFileLogging)
            {
                string directory = Path.GetDirectoryName(Path.GetFullPath(_logFilePath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }
        }

        public void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }

        public void LogWarning(string message)
        {
            WriteLog("WARNING", message);
        }

        public void LogError(string message, Exception ex = null)
        {
            string logMessage = message;
            if (ex != null)
            {
                logMessage += $"\nИсключение: {ex.Message}";
                logMessage += $"\nСтек вызовов: {ex.StackTrace}";
                if (ex.InnerException != null)
                {
                    logMessage += $"\nВнутреннее исключение: {ex.InnerException.Message}";
                }
            }
            WriteLog("ERROR", logMessage);
        }

        private void WriteLog(string level, string message)
        {
            // Создаем запись лога с меткой времени
            string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {level}: {message}";

            // Всегда пишем в Debug
            Debug.WriteLine(logEntry);

            // Пишем в файл, если включено
            if (_enableFileLogging)
            {
                lock (_lockObject)
                {
                    try
                    {
                        File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Не удалось записать в файл лога: {ex.Message}");
                    }
                }
            }
        }
    }
}