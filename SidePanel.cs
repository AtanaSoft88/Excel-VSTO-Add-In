using excelAddInTest.Dto;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

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
            try
            {
                // Invoke our isolated database initializer service layer
                DataResult result = await DataInitializer.LoadPropertyPricesAsync();

                if (!result.IsSuccess)
                {
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
                        priceData[i, 0] = 0; // Gracefully handle invalid format data entries
                }

                // Locate the active spreadsheet instance safely via global reference context
                Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                // Write standard descriptive column header text in row 1
                activeSheet.get_Range("A1").Value2 = "Property Price in €";

                // Setup dynamic target row selection grid boundaries from cell A2 downward
                Excel.Range targetRange = activeSheet.get_Range("A2", "A" + (rowsCount + 1));

                // Process structural in-memory object block array write straight into Excel sheet
                targetRange.Value2 = priceData;

                MessageBox.Show(result.Message, "Data Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to complete database operation: " + ex.Message, "I/O Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- STEP 2: CALCULATE COMMISSION WITH HIGHLY OPTIMIZED IN-MEMORY CACHE ---
        private void calculateBtn_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                // Dynamically evaluate spreadsheet depth based on cell content validation
                Excel.Range lastCell = activeSheet.Cells[activeSheet.Rows.Count, "A"].End(Excel.XlDirection.xlUp);
                int lastRow = lastCell.Row;

                if (lastRow < 2)
                {
                    MessageBox.Show("No active dataset found in Column A. Please initialize data first.", "Data Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // High recall data chunk reading structure prevents Excel thread execution lag
                Excel.Range priceRange = activeSheet.get_Range("A2", "A" + lastRow);
                object[,] priceValues = (object[,])priceRange.Value2;

                object[,] commissionResults = new object[lastRow - 1, 1];

                // Construct temporary application dictionary buffer to hold mathematical lookups
                Dictionary<double, double> commissionCache = new Dictionary<double, double>();

                for (int i = 1; i <= priceValues.GetLength(0); i++)
                {
                    if (priceValues[i, 1] != null && double.TryParse(priceValues[i, 1].ToString(), out double price))
                    {
                        // Check local cache dictionary before running mathematical operations
                        if (commissionCache.ContainsKey(price))
                        {
                            commissionResults[i - 1, 0] = commissionCache[price];
                        }
                        else
                        {
                            // Calculate 3% commission on cache miss
                            double commission = price * 0.03;

                            // Store the result in cache for subsequent duplicate rows
                            commissionCache[price] = commission;
                            commissionResults[i - 1, 0] = commission;
                        }
                    }
                    else
                    {
                        commissionResults[i - 1, 0] = 0;
                    }
                }

                // Place corresponding target business metric header value label
                activeSheet.get_Range("B1").Value2 = "Commission (3%)";

                // Flush result data collection out to spreadsheet workspace grid layout
                Excel.Range resultRange = activeSheet.get_Range("B2", "B" + lastRow);
                resultRange.Value2 = commissionResults;

                MessageBox.Show($"Successfully processed [{lastRow - 1}] entries using In-Memory optimization!", "Calculation Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An unexpected error occurred during bulk calculation: " + ex.Message, "Runtime Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- MANAGEMENT OPERATIONS: CLEAR CALCULATED BUSINESS ENTRIES ---
        private void clearBtn_Click(object sender, EventArgs e)
        {
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

                // Clear structural cell value elements while explicitly preserving workspace borders
                Excel.Range rangeToClear = activeSheet.get_Range("B2", "B" + lastRow);
                rangeToClear.ClearContents();

                MessageBox.Show("Calculated commission data was successfully removed.", "Data Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to clean up target spreadsheet cells: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
