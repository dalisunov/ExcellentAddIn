using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellentAddIn.Objects
{
    public class Pictures
    {
        public Excel.Pictures ExcelPictures { get; private set; }

        public Pictures(Excel.Pictures pictures)
        {
            ExcelPictures = pictures;
        }

        // Методы для работы с картинками в Excel
    }
}
