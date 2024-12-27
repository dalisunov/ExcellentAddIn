using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Events
{
    /// <summary>
    /// События на уровне приложения Excel
    /// </summary>
    public class ApplicationEvents
    {
        private Excel.Application _application;
        private Excel.AppEvents_Event _appEvents;

        public ApplicationEvents(Excel.Application application)
        {
            _application = application;
            _appEvents = (Excel.AppEvents_Event)application;
        }

        /// <summary>
        /// Подписка на события приложения (NewWorkbook, WorkbookOpen и т.д.)
        /// </summary>
        public void Subscribe()
        {
            _appEvents.NewWorkbook += Application_NewWorkbook;
            _appEvents.WorkbookOpen += Application_WorkbookOpen;
            // Можно подписаться на другие интересующие события
        }

        /// <summary>
        /// Отписка от событий, если необходимо
        /// </summary>
        public void Unsubscribe()
        {
            _appEvents.NewWorkbook -= Application_NewWorkbook;
            _appEvents.WorkbookOpen -= Application_WorkbookOpen;
        }

        private void Application_NewWorkbook(Excel.Workbook wb)
        {
            // Логика при создании новой книги
            Console.WriteLine($"New workbook created: {wb.Name}");
        }

        private void Application_WorkbookOpen(Excel.Workbook wb)
        {
            // Логика при открытии книги
            Console.WriteLine($"Workbook opened: {wb.Name}");
        }
    }
}
