using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace excelAddInTest
{
    public partial class ThisAddIn
    {
        // 1. ДЕФИНИРАНЕ: Това е свойството, което MyRibbon търси и не намираше в момента!
        public Microsoft.Office.Tools.CustomTaskPane MyCustomPane { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 2. Initialization of Side Panel
            SidePanel sidePanelUi = new SidePanel();
           
            // 3. Attaching Side Panel to Excel
            MyCustomPane = this.CustomTaskPanes.Add(sidePanelUi, "SaaS Real Estate Analytics");

            // 4. КОНФИГУРАЦИЯ: Задаваме му ширина и позиция отдясно
            MyCustomPane.Width = 320;
            MyCustomPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
