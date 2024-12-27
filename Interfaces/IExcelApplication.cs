using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Interfaces
{
    /// <summary>
    /// Интерфейс, предоставляющий доступ к объекту Excel.Application
    /// </summary>
    public interface IExcelApplication
    {
        /// <summary>
        /// Объект Excel.Application
        /// </summary>
        Excel.Application Application { get; }
    }
}
