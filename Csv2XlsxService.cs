using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.Versioning;
using System.ServiceProcess;
using System.Threading;
using Microsoft.Extensions.Configuration;
using Serilog;
using Serilog.Events;

namespace Csv2Xlsx
{
    [SupportedOSPlatform("windows")]
    public class Csv2XlsxService : ServiceBase
    {
        private FileSystemWatcher? watcher;
        private IConfiguration? configuration;
        private string? templatePath;
        private string? csvFolderPath;
        private string? mappingPath;
        private string? outputFolderPath;
        private bool isRunning = false;
        private CancellationTokenSource? cancellationTokenSource;
        private ILogger? logger;

        public Csv2XlsxService()
        {
            ServiceName = "Csv2XlsxService";
            CanStop = true;
            CanPauseAndContinue = false;
            AutoLog = true;
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                // Register encoding provider for non-Unicode encodings (like windows-1255)
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                // Load configuration from appsettings.json
                configuration = new ConfigurationBuilder()
                    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                // Setup Serilog
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "csv2xlsx-.log");
                
                // Ensure logs directory exists
                Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs"));
                
                // Configure Serilog with rolling file
                Log.Logger = new LoggerConfiguration()
                    .ReadFrom.Configuration(configuration)
                    .MinimumLevel.Debug()
                    .WriteTo.File(
                        path: logPath,
                        rollingInterval: RollingInterval.Day,
                        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                        fileSizeLimitBytes: 10485760, // 10MB
                        retainedFileCountLimit: 31,    // Keep a month of logs
                        rollOnFileSizeLimit: true)
                    .CreateLogger();
                
                logger = Log.Logger;
                logger.Information("Service starting up");

                // Get file paths from configuration and convert to absolute paths
                string basePath = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = configuration["FilePaths:TemplatePath"] ?? "ExcelTemplate.xlsx";
                string csvPath = configuration["FilePaths:CsvFolderPath"] ?? "InputCsv";
                string mapPath = configuration["FilePaths:MappingPath"] ?? "mapping.json";
                string outPath = configuration["FilePaths:OutputFolderPath"] ?? "OutputExcel";
                
                templatePath = Path.GetFullPath(Path.Combine(basePath, tempPath));
                csvFolderPath = Path.GetFullPath(Path.Combine(basePath, csvPath));
                mappingPath = Path.GetFullPath(Path.Combine(basePath, mapPath));
                outputFolderPath = Path.GetFullPath(Path.Combine(basePath, outPath));

                // Log paths
                logger.Information("Starting Csv2Xlsx service with CSV folder: {CsvFolder}", csvFolderPath);
                logger.Information("Output folder: {OutputFolder}", outputFolderPath);

                // Ensure the input and output folders exist
                if (csvFolderPath != null) Program.EnsureFolderExists(csvFolderPath);
                if (outputFolderPath != null) Program.EnsureFolderExists(outputFolderPath);

                // Initialize cancellation token
                cancellationTokenSource = new CancellationTokenSource();

                // Start monitoring in a separate thread
                isRunning = true;
                Thread monitoringThread = new Thread(() => StartMonitoring(cancellationTokenSource.Token));
                monitoringThread.IsBackground = true;
                monitoringThread.Start();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error starting service");
                throw;
            }
        }

        protected override void OnStop()
        {
            try
            {
                logger?.Information("Stopping Csv2Xlsx service");
                
                // Signal the monitoring thread to stop
                isRunning = false;
                cancellationTokenSource?.Cancel();
                
                // Clean up resources
                if (watcher != null)
                {
                    watcher.EnableRaisingEvents = false;
                    watcher.Dispose();
                }
                
                // Flush and close the log
                Log.CloseAndFlush();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error stopping service");
            }
        }

        private void StartMonitoring(CancellationToken cancellationToken)
        {
            try
            {
                if (csvFolderPath == null || templatePath == null || mappingPath == null || outputFolderPath == null || configuration == null)
                {
                    logger?.Error("One or more required paths are null");
                    return;
                }

                // Process any existing files first
                Program.ProcessExistingFiles(csvFolderPath, templatePath, mappingPath, outputFolderPath, configuration);

                // Set up file system watcher
                watcher = new FileSystemWatcher(csvFolderPath)
                {
                    Filter = "*.csv", // Only watch for CSV files
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.LastWrite
                };

                // Event triggered when a new file is created in the folder
                watcher.Created += (sender, e) =>
                {
                    logger?.Information("New file detected: {FileName}", e.Name);
                    
                    // Process the file in a separate thread
                    Thread processThread = new Thread(() =>
                    {
                        // Wait briefly to ensure file is completely written
                        Thread.Sleep(500);
                        if (!cancellationToken.IsCancellationRequested)
                        {
                            try
                            {
                                Program.ProcessCsvFile(e.FullPath, templatePath, mappingPath, outputFolderPath, configuration);
                                logger?.Information("Successfully processed file: {FileName}", e.Name);
                            }
                            catch (Exception ex)
                            {
                                logger?.Error(ex, "Error processing file {FileName}", e.Name);
                            }
                        }
                    });
                    
                    processThread.IsBackground = true;
                    processThread.Start();
                };

                // Start monitoring the folder
                watcher.EnableRaisingEvents = true;
                logger?.Information("Watching folder: {CsvFolder} for new CSV files", csvFolderPath);

                // Keep the service running until stop is requested
                while (isRunning && !cancellationToken.IsCancellationRequested)
                {
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                logger?.Error(ex, "Error in monitoring thread");
            }
        }
    }
}