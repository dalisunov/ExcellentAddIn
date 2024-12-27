using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Objects
{
    public class Charts
    {
        public Excel.ChartObjects ExcelChartObjects { get; private set; }

        public Charts(Excel.ChartObjects chartObjects)
        {
            ExcelChartObjects = chartObjects;
        }

        // Методы для работы с диаграммами
    }
}
