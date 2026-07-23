using System;
using System.Collections.Generic;
using System.Windows.Forms;
using excelAddInTest.Data;
using excelAddInTest.Dto;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace excelAddInTest
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void calcCommissionBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 1. Get access to the currently active Excel worksheet
                Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                // 2. Find the last row with data in Column A (where prices are located)
                Excel.Range lastCell = activeSheet.Cells[activeSheet.Rows.Count, "A"].End(Excel.XlDirection.xlUp);
                int lastRow = lastCell.Row;

                // If no data is entered
                if (lastRow < 2)
                {
                    MessageBox.Show("Please, enter prices in column А (starting from row 2)!");
                    return;
                }

                // 3. BULK ROW OPTIMIZATION (In-Memory Processing):
                // Retrieve the entire range of data at once into a 2D array in C# memory
                Excel.Range priceRange = activeSheet.get_Range("A2", "A" + lastRow);
                object[,] priceValues = (object[,])priceRange.Value2;

                // Prepare an empty array to dump all results into Column B in a single operation
                object[,] commissionResults = new object[lastRow - 1, 1];

                // LOCAL CACHE: Dictionary that stores previously calculated commission values
                Dictionary<double, double> commissionCache = new Dictionary<double, double>();

                // 4. Run a fast in-memory loop through all extracted rows
                for (int i = 1; i <= priceValues.GetLength(0); i++)
                {
                    if (priceValues[i, 1] != null && double.TryParse(priceValues[i, 1].ToString(), out double price))
                    {
                        // Cache lookup - if this price has already been processed, fetch it directly
                        if (commissionCache.ContainsKey(price))
                        {
                            commissionResults[i - 1, 0] = commissionCache[price];
                        }
                        else
                        {
                            // If missing from cache, calculate the 3% commission
                            double commission = price * 0.03;

                            // Store it in cache for subsequent rows with the same price value
                            commissionCache[price] = commission;
                            commissionResults[i - 1, 0] = commission;
                        }
                    }
                    else
                    {
                        commissionResults[i - 1, 0] = 0; // Handle empty cells safely
                    }
                }

                // 5. Write ALL results into Column B simultaneously using a single Excel Interop operation
                Excel.Range resultRange = activeSheet.get_Range("B2", "B" + lastRow);
                resultRange.Value2 = commissionResults;

                MessageBox.Show($"Successfully handled [{lastRow - 1}] rows!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred during Excel data handling: " + ex.Message);
            }
        }

        private void clearBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 1. Get access to the currently active Excel worksheet
                Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                // 2. Find the last populated row in Column B
                Excel.Range lastCell = activeSheet.Cells[activeSheet.Rows.Count, "B"].End(Excel.XlDirection.xlUp);
                int lastRow = lastCell.Row;

                // If there is no data below the header on row 1, there is nothing to clear
                if (lastRow < 2)
                {
                    MessageBox.Show("Column B is already empty!");
                    return;
                }

                // 3. Define the range from B2 to the last row and clear the values at once
                Excel.Range rangeToClear = activeSheet.get_Range("B2", "B" + lastRow);
                rangeToClear.ClearContents(); // Deletes data values only, preserving styles (colors, borders)

                MessageBox.Show("Column B data was successfully cleared!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during data clearing: " + ex.Message);
            }
        }

        private void loadFromDbBtn_Click(object sender, RibbonControlEventArgs e)
        {
            // Invoke our isolated data initialization service class
            DataResult result = DataInitializer.LoadPropertyPrices();

            // If an error occurs or the database is empty, display the message and halt execution
            if (!result.IsSuccess)
            {
                MessageBox.Show(result.Message);
                return;
            }

            // Processing logic and Excel writing remain here as they belong to UI responsibilities
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

            // Display the dynamic success message returned from the initializer service
            MessageBox.Show(result.Message);
        }
    }
}
