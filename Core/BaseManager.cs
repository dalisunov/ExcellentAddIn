using ExcellentAddIn.Interfaces;

namespace ExcellentAddIn.Core
{
    /// <summary>
    /// Базовый класс, от которого могут наследоваться другие менеджеры.
    /// Обеспечивает доступ к ILogger и IErrorHandler
    /// </summary>
    public abstract class BaseManager
    {
        protected readonly ILogger _logger;
        protected readonly IErrorHandler _errorHandler;

        protected BaseManager(ILogger logger, IErrorHandler errorHandler)
        {
            _logger = logger;
            _errorHandler = errorHandler;
        }

        /// <summary>
        /// Пример метода, который может вызываться в производных классах
        /// для логирования или обработки ошибок
        /// </summary>
        protected void SafeExecute(System.Action action)
        {
            try
            {
                action();
            }
            catch (System.Exception ex)
            {
                _errorHandler.HandleError(ex);
            }
        }
    }
}
