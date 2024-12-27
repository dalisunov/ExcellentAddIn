using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Interfaces
{
    public interface IWorkbookManager
    {
        Excel.Workbook GetActiveWorkbook();
        Excel.Workbook OpenWorkbook(string path);
        void CloseWorkbook(Excel.Workbook workbook, bool saveChanges = false);
        // Другие методы, связанные с управлением книгами
    }
}
