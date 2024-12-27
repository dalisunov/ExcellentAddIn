using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Events
{
    /// <summary>
    /// События и операции, связанные с диапазонами Excel
    /// </summary>
    public class RangeEvents
    {
        private Excel.Range _range;

        public RangeEvents(Excel.Range range)
        {
            _range = range;
        }

        public void TrackChange()
        {
            // Пример: если необходимо отследить изменение в конкретном диапазоне
            Console.WriteLine($"Range '{_range.Address}' has changed or needs processing.");
        }

        // Могут быть дополнительные методы, например:
        // - Validating data in range
        // - Format changes
        // - Etc.
    }
}
