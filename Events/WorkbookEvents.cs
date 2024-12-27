using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Events
{
    /// <summary>
    /// События на уровне конкретной Excel-книги
    /// </summary>
    public class WorkbookEvents
    {
        private Excel.Workbook _workbook;

        public WorkbookEvents(Excel.Workbook workbook)
        {
            _workbook = workbook;
            // Подключение к событиям объекта Workbook делается через интерфейсы событий.
            // Можно использовать IDocEventBinding или кастомный approach.
        }

        public void BeforeClose()
        {
            // Пример: подписка на закрытие
            Console.WriteLine($"Workbook '{_workbook.Name}' is about to close.");
        }

        // Другие события Workbook, например:
        public void BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            Console.WriteLine($"Workbook '{_workbook.Name}' is being saved.");
        }

        // И т.д.
    }
}
