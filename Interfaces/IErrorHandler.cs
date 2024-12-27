using System;

namespace ExcellentAddIn.Interfaces
{
    /// <summary>
    /// Определяет контракт для централизованной обработки ошибок в приложении.
    /// Обрабатывает исключения и определяет соответствующие действия.
    /// </summary>
    public interface IErrorHandler
    {
        /// <summary>
        /// Обрабатывает исключение путем его логирования и выполнения соответствующих действий
        /// </summary>
        /// <param name="ex">Исключение для обработки</param>
        void HandleError(Exception ex);

        /// <summary>
        /// Обрабатывает исключение с дополнительной контекстной информацией
        /// </summary>
        /// <param name="ex">Исключение для обработки</param>
        /// <param name="contextMessage">Дополнительный контекст о том, где/почему произошла ошибка</param>
        void HandleError(Exception ex, string contextMessage);
    }
}