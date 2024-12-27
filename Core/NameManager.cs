using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Core
{
    public class NameManager
    {
        // Пример класса для работы с именованными диапазонами (Named Ranges)
        public Excel.Name CreateNamedRange(Excel.Worksheet sheet, string rangeName, string address)
        {
            // Реализация
            return sheet.Names.Add(Name: rangeName, RefersTo: address);
        }

        public void DeleteNamedRange(Excel.Worksheet sheet, string rangeName)
        {
            // Реализация
            foreach (Excel.Name name in sheet.Names)
            {
                if (name.Name == rangeName)
                {
                    name.Delete();
                    break;
                }
            }
        }
    }
}
