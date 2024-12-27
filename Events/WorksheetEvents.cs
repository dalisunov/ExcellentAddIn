using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Events
{
    /// <summary>
    /// События на уровне листа (Worksheet)
    /// </summary>
    public class WorksheetEvents
    {
        private Excel.Worksheet _worksheet;

        public WorksheetEvents(Excel.Worksheet worksheet)
        {
            _worksheet = worksheet;
            // Аналогично WorkbookEvents, 
            // но здесь логика привязана к конкретному листу
        }

        public void OnActivate()
        {
            // Вызывается при активации листа
            Console.WriteLine($"Worksheet '{_worksheet.Name}' is activated.");
        }

        public void OnChange(Excel.Range target)
        {
            // Вызывается при изменении диапазона target
            Console.WriteLine($"Cells changed in worksheet '{_worksheet.Name}'. Changed range: {target.Address}");
        }
    }
}
