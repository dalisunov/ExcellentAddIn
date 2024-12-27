using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Objects
{
    public class PivotTables
    {
        public Excel.PivotTables ExcelPivotTables { get; private set; }

        public PivotTables(Excel.PivotTables pivotTables)
        {
            ExcelPivotTables = pivotTables;
        }

        // Методы для работы с сводными таблицами
    }
}
