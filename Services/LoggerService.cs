using Microsoft.ApplicationInsights.Extensibility;
using Serilog;
using System;
using System.Configuration;
using System.Windows.Forms;

namespace excelAddInTest.Services
{
    public static class LoggerService
    {
        public static void InitializeLogger()
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

                // Fetch the hidden connection string from local App.config appSettings section
                string connectionString = ConfigurationManager.AppSettings["AzureConnectionString"];

                if (string.IsNullOrEmpty(connectionString))
                {
                    // Fallback to a completely universal inline logger (no extra NuGet sinks required)
                    Log.Logger = new LoggerConfiguration()
                        .MinimumLevel.Debug()
                        .CreateLogger();
                    return;
                }

                // Highly optimized Azure Application Insights SDK structure 
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Debug()
                    .WriteTo.ApplicationInsights(
                        new TelemetryConfiguration { ConnectionString = connectionString },
                        TelemetryConverter.Traces)
                    .CreateLogger();

                Log.Information("Excel VSTO Add-In telemetry initialization successful via secure config.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Critical Error during Telemetry System Init: " + ex.Message, "System Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void LogError(string message, Exception ex, string buttonName)
        {
            Log.Error(ex, "EXCEPTION captured in UI element [{ButtonName}]: {Message}", buttonName, message);
        }

        public static void LogAction(string message, string buttonName)
        {
            Log.Information("USER ACTION triggered in [{ButtonName}]: {Message}", buttonName, message);
        }

        public static void ShutdownLogger()
        {
            Log.CloseAndFlush(); // Block execution thread briefly until all cloud log packets are fully shipped
        }
    }
}




//using System;
//using System.Configuration; // Required to read App.config securely
//using System.Windows.Forms;
//using Serilog;
//using Microsoft.ApplicationInsights; // Задължително за TelemetryClient
//using Microsoft.ApplicationInsights.Extensibility;

//namespace excelAddInTest.Services
//{
//    public static class LoggerService
//    {
//        // Създаваме жив TelemetryClient директно от Azure SDK
//        private static TelemetryClient _telemetryClient;

//        public static void InitializeLogger()
//        {
//            try
//            {                
//                // ФОРСИРАМЕ WINDOWS ДА ИЗПОЛЗВА МОДЕРЕН TLS 1.2 ЗА СВЪРЗВАНЕ С ОБЛАКА НА AZURE
//                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

//                string connectionString = ConfigurationManager.AppSettings["AzureConnectionString"];

//                if (string.IsNullOrEmpty(connectionString))
//                {
//                    Log.Logger = new LoggerConfiguration()
//                        .MinimumLevel.Debug()
//                        .CreateLogger();
//                    return;
//                }

//                // 1. Инициализираме живата конфигурация за Azure
//                var config = new TelemetryConfiguration { ConnectionString = connectionString };
//                _telemetryClient = new TelemetryClient(config);

//                // 2. Свързваме Serilog към същата конфигурация
//                Log.Logger = new LoggerConfiguration()
//                    .MinimumLevel.Debug()
//                    .WriteTo.ApplicationInsights(config, TelemetryConverter.Traces)
//                    .CreateLogger();

//                Log.Information("Excel VSTO Add-In telemetry initialization successful via secure client.");
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show("Critical Error during Telemetry System Init: " + ex.Message, "System Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
//            }
//        }

//        public static void LogError(string message, Exception ex, string buttonName)
//        {
//            // 1. Записваме локално през Serilog
//            Log.Error(ex, "EXCEPTION captured in UI element [{ButtonName}]: {Message}", buttonName, message);            

//            // 2. ФОРСИРАМЕ ИЗСТРЕЛВАНЕТО В РЕАЛНО ВРЕМЕ КЪМ LIVE METRICS ЧРЕЗ TELEMETRY CLIENT
//            if (_telemetryClient != null)
//            {
//                if (ex != null)
//                {
//                    _telemetryClient.TrackException(ex);
//                }
//                else
//                {
//                    _telemetryClient.TrackTrace($"[ERROR] Element: {buttonName} - {message}");
//                }
//                _telemetryClient.Flush(); // Насиствено изпращане веднага без чакане!                
//            }
//        }

//        public static void LogAction(string message, string buttonName)
//        {
//            Log.Information("USER ACTION triggered in [{ButtonName}]: {Message}", buttonName, message);

//            // Форсираме и потребителските действия в реално време
//            if (_telemetryClient != null)
//            {
//                _telemetryClient.TrackTrace($"[ACTION] Element: {buttonName} - {message}");
//                _telemetryClient.Flush();
//            }
//        }

//        public static void ShutdownLogger()
//        {
//            if (_telemetryClient != null)
//            {
//                _telemetryClient.Flush();
//            }
//            Log.CloseAndFlush();
//        }
//    }
//}
