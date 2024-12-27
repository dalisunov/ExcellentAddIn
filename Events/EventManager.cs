using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Events
{
    /// <summary>
    /// Центральный менеджер, orchestrator для подписки/отписки от событий
    /// </summary>
    public class EventManager
    {
        private readonly ApplicationEvents _applicationEvents;

        public EventManager(Excel.Application application)
        {
            // Инициализируем классы для подписки на события приложения
            _applicationEvents = new ApplicationEvents(application);
        }

        public void StartListening()
        {
            // Запускаем подписку на события
            _applicationEvents.Subscribe();
            // При необходимости подключаемся к WorkbookEvents, WorksheetEvents, RangeEvents
        }

        public void StopListening()
        {
            // Отписываемся от событий
            _applicationEvents.Unsubscribe();
        }
    }
}
