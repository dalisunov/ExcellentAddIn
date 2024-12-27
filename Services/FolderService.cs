using System.IO;

namespace ExcellentAddIn.Services
{
    public class FolderService
    {
        public bool CreateFolder(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                return true;
            }
            return false;
        }
    }
}
