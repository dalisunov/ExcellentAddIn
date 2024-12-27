using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Interfaces
{
    public interface IWorksheetManager
    {
        Excel.Worksheet GetActiveWorksheet();
        Excel.Worksheet AddWorksheet(string name = null);
        // Другие методы, связанные с управлением листами
    }
}
