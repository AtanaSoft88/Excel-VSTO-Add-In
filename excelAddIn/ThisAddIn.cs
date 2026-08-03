using System;
using Excel = Microsoft.Office.Interop.Excel;
using excelAddInTest.Services; // Задължително добавяме това, за да намери логъра в папката Services!

namespace excelAddInTest
{
    public partial class ThisAddIn
    {
        // 1. Definition of the custom task pane
        public Microsoft.Office.Tools.CustomTaskPane MyCustomPane { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // ==========================================
            // 1. Starting logger system (AZURE MONITORS)
            LoggerService.InitializeLogger();
            // ==========================================

            // 2. Initialization of the curved Side Panel UI
            SidePanel sidePanelUi = new SidePanel();

            // 3. Attaching the Side Panel to Excel workspace
            MyCustomPane = this.CustomTaskPanes.Add(sidePanelUi, "SaaS Real Estate Analytics");

            // 4. Configure Panel location - docked to the right side
            MyCustomPane.Width = 320;
            MyCustomPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // ==========================================
            // Guarantee all logging packages are sent to Azure before closing.
            LoggerService.ShutdownLogger();
            // ==========================================
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
