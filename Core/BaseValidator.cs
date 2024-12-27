using System;
using ExcellentAddIn.Interfaces;

namespace ExcellentAddIn.Core
{
    /// <summary>
    /// Базовый класс для всех валидаторов, предоставляющий общую функциональность
    /// </summary>
    public abstract class BaseValidator
    {
        protected readonly ILogger _logger;
        protected readonly IErrorHandler _errorHandler;

        protected BaseValidator(ILogger logger, IErrorHandler errorHandler)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _errorHandler = errorHandler ?? throw new ArgumentNullException(nameof(errorHandler));
        }

        /// <summary>
        /// Выполняет валидацию и логирует результат
        /// </summary>
        protected bool ValidateWithLogging(Func<bool> validationFunc, string validationMessage)
        {
            try
            {
                bool result = validationFunc();
                if (!result)
                {
                    _logger.LogWarning($"Ошибка валидации: {validationMessage}");
                }
                return result;
            }
            catch (Exception ex)
            {
                _errorHandler.HandleError(ex, $"Ошибка при выполнении валидации: {validationMessage}");
                return false;
            }
        }
    }
}