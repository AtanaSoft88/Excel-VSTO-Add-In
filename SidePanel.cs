using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using excelAddInTest.Services; // ЗАДЪЛЖИТЕЛНО: Включваме логъра от неговата папка!
using excelAddInTest.Dto;

namespace excelAddInTest
{
    public partial class SidePanel : UserControl
    {
        public SidePanel()
        {
            InitializeComponent();
        }

        // --- STEP 1: LOAD PROPERTY PRICES FROM LOCAL MOCK DATABASE ---
        private async void loadBtn_Click(object sender, EventArgs e)
        {
            // TELEMETRY: Record that user initialized data load sequence
            LoggerService.LogAction("User triggered remote database synchronization process.", nameof(loadBtn));

            try
            {
                loadBtn.Text = "Loading from API...";
                loadBtn.Enabled = false;

                DataResult result = await DataInitializer.LoadPropertyPricesAsync();

                loadBtn.Text = "Load from DB";
                loadBtn.Enabled = true;

                if (!result.IsSuccess)
                {
                    LoggerService.LogError($"API Sync failed. Message returned: {result.Message}", null, nameof(loadBtn));
                    MessageBox.Show(result.Message, "Database Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int rowsCount = result.PropertyPrices.Length;
                object[,] priceData = new object[rowsCount, 1];

                for (int i = 0; i < rowsCount; i++)
                {
                    if (double.TryParse(result.PropertyPrices[i], out double price))
                        priceData[i, 0] = price;
                    else
                        priceData[i, 0] = 0;
                }

                Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                activeSheet.get_Range("A1").Value2 = "Property Price in €";

                Excel.Range targetRange = activeSheet.get_Range("A2", "A" + (rowsCount + 1));
                targetRange.Value2 = priceData;

                LoggerService.LogAction($"Successfully rendered [{rowsCount}] property rows directly to worksheet active layout.", nameof(loadBtn));
                MessageBox.Show(result.Message, "Data Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                loadBtn.Text = "Load from DB";
                loadBtn.Enabled = true;

                // OBSERVABILITY: Shifting full thread exception metrics up to Azure Cloud Logs Workspace
                LoggerService.LogError("Critical I/O connection failure occurred during spreadsheet loading.", ex, nameof(loadBtn));
                MessageBox.Show("Failed to complete database operation: " + ex.Message, "I/O Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- STEP 2: CALCULATE COMMISSION WITH HIGHLY OPTIMIZED IN-MEMORY CACHE ---
        private void calculateBtn_Click(object sender, EventArgs e)
        {
            // TELEMETRY: Trace user intent
            LoggerService.LogAction("User initiated optimization engine for commission bulk calculation.", nameof(calculateBtn));

            try
            {
                Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                Excel.Range lastCell = activeSheet.Cells[activeSheet.Rows.Count, "A"].End(Excel.XlDirection.xlUp);
                int lastRow = lastCell.Row;

                if (lastRow < 2)
                {
                    // OBSERVABILITY: Log specific structural warning data before dropping out
                    LoggerService.LogError("Calculation halted. Reason: Column A data context was empty.", null, nameof(calculateBtn));
                    MessageBox.Show("No active dataset found in Column A. Please initialize data first.", "Data Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Excel.Range priceRange = activeSheet.get_Range("A2", "A" + lastRow);
                object[,] priceValues = (object[,])priceRange.Value2;

                object[,] commissionResults = new object[lastRow - 1, 1];
                Dictionary<double, double> commissionCache = new Dictionary<double, double>();

                for (int i = 1; i <= priceValues.GetLength(0); i++)
                {
                    if (priceValues[i, 1] != null && double.TryParse(priceValues[i, 1].ToString(), out double price))
                    {
                        if (commissionCache.ContainsKey(price))
                        {
                            commissionResults[i - 1, 0] = commissionCache[price];
                        }
                        else
                        {
                            double commission = price * 0.03;
                            commissionCache[price] = commission;
                            commissionResults[i - 1, 0] = commission;
                        }
                    }
                    else
                    {
                        commissionResults[i - 1, 0] = 0;
                    }
                }

                activeSheet.get_Range("B1").Value2 = "Commission (3%)";
                Excel.Range resultRange = activeSheet.get_Range("B2", "B" + lastRow);
                resultRange.Value2 = commissionResults;

                LoggerService.LogAction($"Bulk pricing evaluation finished successfully. Handled entries: [{lastRow - 1}].", nameof(calculateBtn));
                MessageBox.Show($"Successfully processed [{lastRow - 1}] entries using In-Memory optimization!", "Calculation Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // OBSERVABILITY: Ship complete system crashed analytics directly to Azure
                LoggerService.LogError("Runtime mathematical matrix calculation collapsed during execution runtime.", ex, nameof(calculateBtn));
                MessageBox.Show("An unexpected error occurred during bulk calculation: " + ex.Message, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- MANAGEMENT OPERATIONS: CLEAR CALCULATED BUSINESS ENTRIES ---
        private void clearBtn_Click(object sender, EventArgs e)
        {
            LoggerService.LogAction("User triggered workspace layout sanitization.", nameof(clearBtn));

            try
            {
                Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                Excel.Range lastCell = activeSheet.Cells[activeSheet.Rows.Count, "B"].End(Excel.XlDirection.xlUp);
                int lastRow = lastCell.Row;

                if (lastRow < 2)
                {
                    MessageBox.Show("Target column range is already empty.", "Clear Operation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Excel.Range rangeToClear = activeSheet.get_Range("B2", "B" + lastRow);
                rangeToClear.ClearContents();

                LoggerService.LogAction("Worksheet targets flushed and wiped clean successfully.", nameof(clearBtn));
                MessageBox.Show("Calculated commission data was successfully removed.", "Data Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LoggerService.LogError("Exception captured while attempting to execute cell range execution wipe.", ex, nameof(clearBtn));
                MessageBox.Show("Failed to clean up target spreadsheet cells: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
