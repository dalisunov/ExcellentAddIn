using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Objects
{
    public class Worksheets
    {
        public Excel.Sheets ExcelWorksheets { get; private set; }

        public Worksheets(Excel.Sheets sheets)
        {
            ExcelWorksheets = sheets;
        }

        // Методы для взаимодействия с листами
    }
}
