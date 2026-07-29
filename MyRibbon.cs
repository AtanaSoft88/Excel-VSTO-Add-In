using System;
using Microsoft.Office.Tools.Ribbon;

namespace excelAddInTest
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        // This is the only event remaining in the Ribbon code.
        // It toggles the visibility of your custom SidePanel right inside Excel.
        private void btnTogglePanel_Click(object sender, RibbonControlEventArgs e)
        {
            // Verify if the custom task pane instance is initialized
            if (Globals.ThisAddIn.MyCustomPane != null)
            {
                // Toggle visibility: if open, close it; if hidden, show it
                Globals.ThisAddIn.MyCustomPane.Visible = !Globals.ThisAddIn.MyCustomPane.Visible;
            }
        }
    }
}
