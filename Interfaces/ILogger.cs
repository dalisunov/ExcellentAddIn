using System;

namespace ExcellentAddIn.Interfaces
{
    /// <summary>
    /// Определяет контракт для функционала логирования в приложении.
    /// Поддерживает различные уровни логирования и опциональные детали исключений.
    /// </summary>
    public interface ILogger
    {
        /// <summary>
        /// Логирует информационное сообщение для отслеживания работы приложения
        /// </summary>
        /// <param name="message">Информационное сообщение для записи в лог</param>
        void LogInfo(string message);

        /// <summary>
        /// Логирует предупреждение для потенциально опасных, но некритичных ситуаций
        /// </summary>
        /// <param name="message">Сообщение с предупреждением для записи в лог</param>
        void LogWarning(string message);

        /// <summary>
        /// Логирует сообщение об ошибке с опциональными деталями исключения
        /// </summary>
        /// <param name="message">Сообщение об ошибке для записи в лог</param>
        /// <param name="ex">Опциональный объект исключения с дополнительными деталями</param>
        void LogError(string message, Exception ex = null);
    }
}