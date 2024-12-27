using System;
using System.Linq;
using System.Text.RegularExpressions;
using ExcellentAddIn.Interfaces;

namespace ExcellentAddIn.Core
{
    /// <summary>
    /// Класс для валидации различных типов данных
    /// </summary>
    public class DataValidation : BaseValidator
    {
        public DataValidation(ILogger logger, IErrorHandler errorHandler)
            : base(logger, errorHandler) { }

        /// <summary>
        /// Проверяет, является ли значение числом
        /// </summary>
        public bool IsNumeric(object value)
        {
            return ValidateWithLogging(() =>
                value != null && decimal.TryParse(value.ToString(), out _),
                "Значение должно быть числом");
        }

        /// <summary>
        /// Проверяет формат даты
        /// </summary>
        public bool IsValidDate(string date)
        {
            return ValidateWithLogging(() =>
                DateTime.TryParse(date, out _),
                "Некорректный формат даты");
        }

        /// <summary>
        /// Проверяет формат email
        /// </summary>
        public bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email)) return false;

            return ValidateWithLogging(() =>
            {
                try
                {
                    var regex = new Regex(@"^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$");
                    return regex.IsMatch(email);
                }
                catch
                {
                    return false;
                }
            }, "Некорректный формат email");
        }

        /// <summary>
        /// Проверяет, что строка содержит только буквы и цифры
        /// </summary>
        public bool IsAlphanumeric(string value)
        {
            return ValidateWithLogging(() =>
                !string.IsNullOrEmpty(value) && value.All(char.IsLetterOrDigit),
                "Значение должно содержать только буквы и цифры");
        }
    }
}