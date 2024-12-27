using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace ExcellentAddIn
{
    [ComVisible(true)]
    public class CustomRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        #region IRibbonExtensibility Members
        public string GetCustomUI(string ribbonID)
        {
            // Загружаем CustomRibbon.xml из ресурсов
            return GetResourceText("ExcellentAddIn.Ribbon.CustomRibbon.xml");
        }
        #endregion

        #region Ribbon Callbacks
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnTestButtonClick(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Работает!", "Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Helpers
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            foreach (string resource in resourceNames)
            {
                if (string.Compare(resourceName, resource, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resource)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion
    }
}
