using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Objects
{
    public class Ranges
    {
        public Excel.Range ExcelRange { get; private set; }

        public Ranges(Excel.Range range)
        {
            ExcelRange = range;
        }

        // Методы для работы с конкретным диапазоном
    }
}
