using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Objects
{
    public class ListObjects
    {
        public Excel.ListObjects ExcelListObjects { get; private set; }

        public ListObjects(Excel.ListObjects listObjects)
        {
            ExcelListObjects = listObjects;
        }

        // Методы для работы с ListObjects (таблицами)
    }
}
