using System.IO;

namespace ExcellentAddIn.Services
{
    public class FileService
    {
        public void SaveTextToFile(string path, string content)
        {
            File.WriteAllText(path, content);
        }

        public string ReadTextFromFile(string path)
        {
            return File.Exists(path) ? File.ReadAllText(path) : string.Empty;
        }
    }
}
