using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Format
{
    public class ConditionalFormatting
    {
        public void AddColorScale(Excel.Range range)
        {
            var format = range.FormatConditions.AddColorScale(3);
            // Настройки gradation
        }
    }
}
