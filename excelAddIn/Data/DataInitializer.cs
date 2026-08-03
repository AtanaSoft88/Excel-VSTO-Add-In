using excelAddInTest;
using System;
using System.IO;
using System.Threading.Tasks;
namespace excelAddInTest
{
    public static class DataInitializer
    {
        private static readonly string[] DefaultPrices = { "120000", "145000", "110000", "320000", "95000", "500000", "215000", "180000", "133000", "132000" };
        private const string PropertiesFileName = "properties_db.txt";
        private const string DbEmptyMessage = "Database is empty!";

        // Променяме метода на "public static async Task<DataResult>"
        public static async Task<DataResult> LoadPropertyPricesAsync()
        {
            try
            {
                // Network slowdown situation           
                await Task.Delay(2000);

                string propertyPricesDb = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, PropertiesFileName);

                if (!File.Exists(propertyPricesDb))
                {
                    File.WriteAllLines(propertyPricesDb, DefaultPrices);
                }

                string[] dbLines = File.ReadAllLines(propertyPricesDb);
                int rowsCount = dbLines.Length;

                if (rowsCount == 0)
                {
                    return new DataResult
                    {
                        Message = DbEmptyMessage,
                        PropertyPrices = dbLines,
                        IsSuccess = false
                    };
                }

                string successMessage = $"Successfully loaded [{rowsCount}] properties from remote API!";

                return new DataResult
                {
                    Message = successMessage,
                    PropertyPrices = dbLines,
                    IsSuccess = true
                };
            }
            catch (Exception ex)
            {
                return new DataResult
                {
                    Message = $"Failed to load data: {ex.Message}",
                    PropertyPrices = new string[0],
                    IsSuccess = false
                };
            }
        }
    }
}

