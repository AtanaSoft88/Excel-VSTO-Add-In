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
