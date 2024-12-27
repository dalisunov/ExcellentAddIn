using System;
using ExcellentAddIn.Interfaces;

namespace ExcellentAddIn.Core
{
    /// <summary>
    /// Класс для валидации общих объектов и их свойств
    /// </summary>
    public class ObjectValidation : BaseValidator
    {
        public ObjectValidation(ILogger logger, IErrorHandler errorHandler)
            : base(logger, errorHandler) { }

        /// <summary>
        /// Проверяет, что объект не является null
        /// </summary>
        public bool IsValidObject(object obj, string objectName = "Объект")
        {
            return ValidateWithLogging(
                () => obj != null,
                $"{objectName} не может быть null");
        }

        /// <summary>
        /// Проверяет, что строковое свойство объекта не пустое
        /// </summary>
        public bool HasValidStringProperty(object obj, string propertyName)
        {
            return ValidateWithLogging(() =>
            {
                if (obj == null) return false;
                var property = obj.GetType().GetProperty(propertyName);
                if (property == null) return false;

                var value = property.GetValue(obj) as string;
                return !string.IsNullOrWhiteSpace(value);
            }, $"Свойство {propertyName} должно содержать непустое значение");
        }

        /// <summary>
        /// Проверяет, что числовое свойство находится в допустимом диапазоне
        /// </summary>
        public bool IsNumericPropertyInRange(object obj, string propertyName, decimal minValue, decimal maxValue)
        {
            return ValidateWithLogging(() =>
            {
                if (obj == null) return false;
                var property = obj.GetType().GetProperty(propertyName);
                if (property == null) return false;

                var value = Convert.ToDecimal(property.GetValue(obj));
                return value >= minValue && value <= maxValue;
            }, $"Значение {propertyName} должно быть между {minValue} и {maxValue}");
        }
    }
}