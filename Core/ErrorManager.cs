using System;
using System.IO;
using System.Windows.Forms;
using ExcellentAddIn.Interfaces;

namespace ExcellentAddIn.Core
{
    /// <summary>
    /// Централизует логику обработки ошибок с поддержкой логирования и уведомлений пользователя.
    /// Реализует graceful восстановление после ошибок и понятные пользователю сообщения.
    /// </summary>
    public class ErrorManager : IErrorHandler
    {
        private readonly ILogger _logger;
        private const string СтандартноеСообщениеОбОшибке =
            "Произошла непредвиденная ошибка. Пожалуйста, попробуйте снова или обратитесь в поддержку.";

        public ErrorManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public void HandleError(Exception ex)
        {
            HandleError(ex, null);
        }

        public void HandleError(Exception ex, string contextMessage)
        {
            if (ex == null) return;

            // Логируем ошибку с полными деталями
            _logger.LogError(contextMessage ?? "Произошло необработанное исключение", ex);

            // Определяем понятное пользователю сообщение на основе типа исключения
            string userMessage = ПолучитьПонятноеСообщение(ex);

            // Добавляем контекст, если предоставлен
            if (!string.IsNullOrEmpty(contextMessage))
            {
                userMessage = $"{contextMessage}\n\n{userMessage}";
            }

            // Показываем сообщение пользователю
            ПоказатьОшибкуПользователю(userMessage);
        }

        private string ПолучитьПонятноеСообщение(Exception ex)
        {
            // Преобразуем технические исключения в понятные пользователю сообщения
            if (ex is UnauthorizedAccessException)
            {
                return "У вас нет прав для выполнения этой операции. Пожалуйста, обратитесь к администратору.";
            }
            if (ex is IOException)
            {
                return "Возникла проблема при доступе к файлу или сетевому ресурсу. Пожалуйста, проверьте права доступа и попробуйте снова.";
            }
            if (ex is ArgumentException)
            {
                return "Предоставлены некорректные данные. Пожалуйста, проверьте введенные значения и попробуйте снова.";
            }
            // Добавьте другие типы исключений при необходимости
            return СтандартноеСообщениеОбОшибке;
        }

        private void ПоказатьОшибкуПользователю(string message)
        {
            try
            {
                // Показываем сообщение об ошибке в UI потоке
                if (Application.OpenForms.Count > 0)
                {
                    Form mainForm = Application.OpenForms[0];
                    mainForm.Invoke((MethodInvoker)delegate
                    {
                        MessageBox.Show(
                            mainForm,
                            message,
                            "Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    });
                }
                else
                {
                    MessageBox.Show(
                        message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                // Если показ сообщения об ошибке не удался, хотя бы логируем её
                _logger.LogError("Не удалось показать сообщение об ошибке пользователю", ex);
            }
        }
    }
}