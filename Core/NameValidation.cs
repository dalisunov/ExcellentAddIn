using System;
using System.Text.RegularExpressions;
using ExcellentAddIn.Interfaces;

namespace ExcellentAddIn.Core
{
    /// <summary>
    /// Класс для валидации имён в Excel (листов, диапазонов, etc.)
    /// </summary>
    public class NameValidation : BaseValidator
    {
        // Максимальная длина имени в Excel
        private const int MaxExcelNameLength = 255;

        // Недопустимые символы в именах Excel
        private static readonly char[] InvalidChars = new[] { '\\', '/', '?', '*', '[', ']' };

        public NameValidation(ILogger logger, IErrorHandler errorHandler)
            : base(logger, errorHandler) { }

        /// <summary>
        /// Проверяет допустимость имени листа Excel
        /// </summary>
        public bool IsValidSheetName(string name)
        {
            return ValidateWithLogging(() =>
            {
                if (string.IsNullOrWhiteSpace(name)) return false;
                if (name.Length > MaxExcelNameLength) return false;
                if (name.IndexOfAny(InvalidChars) != -1) return false;

                // Специфичные правила для листов
                return !name.StartsWith("'") && !name.EndsWith("'");
            }, "Недопустимое имя листа");
        }

        /// <summary>
        /// Проверяет допустимость имени именованного диапазона
        /// </summary>
        public bool IsValidRangeName(string name)
        {
            return ValidateWithLogging(() =>
            {
                if (string.IsNullOrWhiteSpace(name)) return false;
                if (name.Length > MaxExcelNameLength) return false;
                if (name.IndexOfAny(InvalidChars) != -1) return false;

                // Правила для именованных диапазонов
                var regex = new Regex("^[a-zA-Z_][a-zA-Z0-9_]*$");
                return regex.IsMatch(name);
            }, "Недопустимое имя диапазона");
        }

        /// <summary>
        /// Проверяет корректность адреса ячейки или диапазона
        /// </summary>
        public bool IsValidRangeAddress(string address)
        {
            return ValidateWithLogging(() =>
            {
                if (string.IsNullOrWhiteSpace(address)) return false;

                // Базовая проверка формата адреса (например, A1, $A$1, A1:B2)
                var regex = new Regex(@"^(\$?[A-Za-z]+\$?\d+)(:\$?[A-Za-z]+\$?\d+)?$");
                return regex.IsMatch(address);
            }, "Некорректный адрес диапазона");
        }

        /// <summary>
        /// Проверяет допустимость имени макроса
        /// </summary>
        public bool IsValidMacroName(string name)
        {
            return ValidateWithLogging(() =>
            {
                if (string.IsNullOrWhiteSpace(name)) return false;
                if (name.Length > MaxExcelNameLength) return false;

                // Правила для имён макросов
                var regex = new Regex("^[a-zA-Z][a-zA-Z0-9_]*$");
                return regex.IsMatch(name);
            }, "Недопустимое имя макроса");
        }
    }
}