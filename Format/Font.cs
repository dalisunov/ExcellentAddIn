using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Format
{
    public class Font
    {
        public void SetFontProperties(Excel.Range range, string fontName = "Calibri", int size = 11)
        {
            range.Font.Name = fontName;
            range.Font.Size = size;
        }
    }
}
