using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Objects
{
    public class Workbooks
    {
        public Excel.Workbooks ExcelWorkbooks { get; private set; }

        public Workbooks(Excel.Workbooks workbooks)
        {
            ExcelWorkbooks = workbooks;
        }

        // Методы для взаимодействия со всеми книгами
    }
}
