using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Format
{
    public class Color
    {
        public void SetBackgroundColor(Excel.Range range, int colorIndex)
        {
            // Пример: использование ColorIndex
            range.Interior.ColorIndex = colorIndex;
        }
    }
}
