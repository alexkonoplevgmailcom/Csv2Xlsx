﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using System.ServiceProcess;
using Serilog;

namespace Csv2Xlsx
{
    partial class Program
    {
        static void Main(string[] args)
        {
            // Check if the application should run as a service
            if (args.Length > 0 && args[0].ToLower() == "--service")
            {
                // Run as a Windows Service
                ServiceBase[] ServicesToRun = new ServiceBase[]
                {
                    new Csv2XlsxService()
                };
                ServiceBase.Run(ServicesToRun);
            }
            else
            {
                // Run as a console application
                
                // Register encoding provider for non-Unicode encodings (like windows-1255)
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                // Load configuration from appsettings.json
                var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                // Setup Serilog for console mode
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "csv2xlsx-.log");
                
                // Ensure logs directory exists
                Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs"));
                
                // Configure Serilog with rolling file and console output
                Log.Logger = new LoggerConfiguration()
                    .ReadFrom.Configuration(configuration)
                    .MinimumLevel.Debug()
                    .WriteTo.Console()
                    .WriteTo.File(
                        path: logPath,
                        rollingInterval: RollingInterval.Day,
                        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                        fileSizeLimitBytes: 10485760, // 10MB
                        retainedFileCountLimit: 31,    // Keep a month of logs
                        rollOnFileSizeLimit: true)
                    .CreateLogger();

                try
                {
                    // Get file paths from configuration and convert to absolute paths
                    string basePath = AppDomain.CurrentDomain.BaseDirectory;
                    string templatePath = Path.GetFullPath(Path.Combine(basePath, configuration["FilePaths:TemplatePath"] ?? "ExcelTemplate.xlsx"));
                    string csvFolderPath = Path.GetFullPath(Path.Combine(basePath, configuration["FilePaths:CsvFolderPath"] ?? "InputCsv"));
                    string mappingPath = Path.GetFullPath(Path.Combine(basePath, configuration["FilePaths:MappingPath"] ?? "mapping.json"));
                    string outputFolderPath = Path.GetFullPath(Path.Combine(basePath, configuration["FilePaths:OutputFolderPath"] ?? "OutputExcel"));

                    Log.Information("Using CSV folder: {CsvFolder}", csvFolderPath);
                    Log.Information("Using output folder: {OutputFolder}", outputFolderPath);
                    Console.WriteLine($"Using CSV folder: {csvFolderPath}");
                    Console.WriteLine($"Using output folder: {outputFolderPath}");

                    // Ensure the input and output folders exist
                    EnsureFolderExists(csvFolderPath);
                    EnsureFolderExists(outputFolderPath);

                    // Start monitoring the folder for new CSV files
                    MonitorFolder(csvFolderPath, templatePath, mappingPath, outputFolderPath, configuration);

                    Console.WriteLine("Monitoring folder for new CSV files. Press Enter to exit...");
                    Console.ReadLine();
                    
                    // Close and flush log when application ends
                    Log.CloseAndFlush();
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Error in main program");
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }
        }

        // Ensures that the specified folder exists, creating it if necessary
        public static void EnsureFolderExists(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                Log.Information("Created missing folder: {FolderPath}", folderPath);
                Console.WriteLine($"Created missing folder: {folderPath}");
            }
        }

        // Monitors the specified folder for new CSV files and processes them
        static void MonitorFolder(string folderPath, string templatePath, string mappingPath, string outputFolderPath, IConfiguration configuration)
        {
            // Process any existing files first
            ProcessExistingFiles(folderPath, templatePath, mappingPath, outputFolderPath, configuration);

            var watcher = new FileSystemWatcher(folderPath)
            {
                Filter = "*.csv", // Only watch for CSV files
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.LastWrite // Watch for file creation and modification
            };

            // Event triggered when a new file is created in the folder
            watcher.Created += (sender, e) =>
            {
                Log.Information("New file detected: {FileName}", e.Name);
                Console.WriteLine($"New file detected: {e.Name}");
                
                // Process the file asynchronously to allow multiple files to be handled simultaneously
                Task.Run(() => 
                {
                    // Wait briefly to ensure file is completely written
                    System.Threading.Thread.Sleep(500);
                    ProcessCsvFile(e.FullPath, templatePath, mappingPath, outputFolderPath, configuration);
                });
            };

            // Start monitoring the folder
            watcher.EnableRaisingEvents = true;
            Log.Information("Watching folder: {FolderPath} for new CSV files", folderPath);
            Console.WriteLine($"Watching folder: {folderPath} for new CSV files");
        }

        // Process existing CSV files in the folder
        public static void ProcessExistingFiles(string folderPath, string templatePath, string mappingPath, string outputFolderPath, IConfiguration configuration)
        {
            try
            {
                // Get all CSV files in the folder
                string[] csvFiles = Directory.GetFiles(folderPath, "*.csv");
                
                if (csvFiles.Length > 0)
                {
                    Log.Information("Found {FileCount} existing CSV files to process", csvFiles.Length);
                    Console.WriteLine($"Found {csvFiles.Length} existing CSV files to process");
                    
                    // Process each file
                    foreach (string csvFile in csvFiles)
                    {
                        string fileName = Path.GetFileName(csvFile);
                        Log.Information("Processing existing file: {FileName}", fileName);
                        Console.WriteLine($"Processing existing file: {fileName}");
                        
                        // Check if the file has already been processed
                        string outputFileName = Path.GetFileNameWithoutExtension(csvFile) + "_output.xlsx";
                        string outputPath = Path.Combine(outputFolderPath, outputFileName);
                        
                        if (!File.Exists(outputPath))
                        {
                            ProcessCsvFile(csvFile, templatePath, mappingPath, outputFolderPath, configuration);
                        }
                        else
                        {
                            Log.Information("Skipping {FileName} - already processed", fileName);
                            Console.WriteLine($"Skipping {fileName} - already processed");
                        }
                    }
                }
                else
                {
                    Log.Information("No existing CSV files found to process");
                    Console.WriteLine("No existing CSV files found to process");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error processing existing files");
                Console.WriteLine($"Error processing existing files: {ex.Message}");
            }
        }

        // Processes a single CSV file and generates an Excel file based on the template and mapping
        public static void ProcessCsvFile(string csvPath, string templatePath, string mappingPath, string outputFolderPath, IConfiguration configuration)
        {
            try
            {
                Log.Information("Processing file: {FilePath}", csvPath);
                Console.WriteLine($"Processing file: {csvPath}");

                // Load the mapping configuration
                var mapping = LoadMapping(mappingPath);

                // Read the CSV data into a list of dictionaries
                // Use the correct configuration path for encoding
                var csvData = ReadCsv(csvPath, configuration["Encoding:CsvEncoding"]);

                // Generate the output file path
                string outputFileName = Path.GetFileNameWithoutExtension(csvPath) + "_output.xlsx";
                string outputPath = Path.Combine(outputFolderPath, outputFileName);

                // Generate the Excel file based on the template and mapping
                GenerateExcel(templatePath, csvData, mapping, outputPath);

                Log.Information("Excel file generated successfully: {OutputPath}", outputPath);
                Log.Information("File processing completed: {FilePath}", csvPath);
                Console.WriteLine($"Excel file generated successfully: {outputPath}");
                Console.WriteLine($"File processing completed: {csvPath}");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during processing
                Log.Error(ex, "Error processing file {FilePath}", csvPath);
                Console.WriteLine($"Error processing file {csvPath}: {ex.Message}");
            }
        }

        // Loads the mapping configuration from the specified JSON file
        static Mapping LoadMapping(string mappingPath)
        {
            string json = File.ReadAllText(mappingPath);
            var mapping = JsonConvert.DeserializeObject<Mapping>(json);
            if (mapping == null)
            {
                Log.Error("Failed to deserialize mapping file: {FilePath}", mappingPath);
                throw new InvalidOperationException($"Failed to deserialize mapping file: {mappingPath}");
            }
            return mapping;
        }

        // Reads the CSV file and converts it into a list of dictionaries
        static List<Dictionary<string, string>> ReadCsv(string csvPath, string? encodingName)
        {
            var csvData = new List<Dictionary<string, string>>();
            
            // Get the encoding from the configuration with a fallback to UTF-8
            System.Text.Encoding encoding;
            try 
            {
                encoding = System.Text.Encoding.GetEncoding(encodingName ?? "utf-8");
            }
            catch (ArgumentException)
            {
                Log.Warning("Encoding '{Encoding}' not found. Using UTF-8 instead", encodingName);
                Console.WriteLine($"Warning: Encoding '{encodingName}' not found. Using UTF-8 instead.");
                encoding = System.Text.Encoding.UTF8;
            }
            
            // Read all lines from the CSV file using the specified encoding
            var lines = File.ReadAllLines(csvPath, encoding);
            
            // The first line contains the column headers
            var headers = lines[0].Split(',');
            
            // Process each line after the header
            foreach (var line in lines.Skip(1))
            {
                var values = line.Split(',');
                var row = new Dictionary<string, string>();
                for (int i = 0; i < Math.Min(headers.Length, values.Length); i++)
                {
                    row[headers[i]] = values[i]; // Map each value to its corresponding header
                }
                csvData.Add(row);
            }
            
            return csvData;
        }

        // Generates an Excel file based on the template, CSV data, and mapping configuration
        static void GenerateExcel(string templatePath, List<Dictionary<string, string>> csvData, Mapping mapping, string outputPath)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook(templatePath))
            {
                // Get the worksheet by name or default to the first worksheet
                var worksheet = !string.IsNullOrEmpty(mapping.SheetName)
                    ? workbook.Worksheet(mapping.SheetName)
                    : workbook.Worksheet(1);

                int startRow = mapping.StartRow; // The starting row for writing data

                // Write data based on column mappings
                foreach (var row in csvData)
                {
                    foreach (var map in mapping.ColumnMappings)
                    {
                        string csvColumn = map.Key;
                        string excelColumn = map.Value;

                        if (row.ContainsKey(csvColumn) && !string.IsNullOrEmpty(row[csvColumn]))
                        {
                            int columnIndex = ExcelColumnToIndex(excelColumn);
                            worksheet.Cell(startRow, columnIndex).Value = row[csvColumn];
                        }
                    }
                    startRow++; // Move to the next row for the next set of data
                }

                // Write data based on specific cell mappings
                foreach (var cellMapping in mapping.CellMappings)
                {
                    string csvColumn = cellMapping.CsvColumn;
                    int csvRow = cellMapping.CsvRow - 1; // Adjust for zero-based index
                    string excelColumn = cellMapping.ExcelColumn;

                    // Determine the Excel row
                    int excelRow;
                    if (cellMapping.ExcelRow.HasValue)
                    {
                        excelRow = cellMapping.ExcelRow.Value; // Use the hardcoded row number
                    }
                    else if (cellMapping.OffsetFromEnd.HasValue)
                    {
                        excelRow = startRow + cellMapping.OffsetFromEnd.Value; // Calculate row based on offset
                    }
                    else
                    {
                        throw new InvalidOperationException("Either ExcelRow or OffsetFromEnd must be specified for a cell mapping.");
                    }

                    if (csvRow >= 0 && csvRow < csvData.Count && csvData[csvRow].ContainsKey(csvColumn))
                    {
                        string value = csvData[csvRow][csvColumn];
                        if (!string.IsNullOrEmpty(value))
                        {
                            int columnIndex = ExcelColumnToIndex(excelColumn);
                            worksheet.Cell(excelRow, columnIndex).Value = value;
                        }
                    }
                }

                // Save the generated Excel file
                workbook.SaveAs(outputPath);
            }
        }

        // Converts an Excel column letter (e.g., "A") to a numeric index (e.g., 1)
        static int ExcelColumnToIndex(string column)
        {
            int index = 0;
            foreach (char c in column.ToUpper())
            {
                index = index * 26 + (c - 'A' + 1);
            }
            return index;
        }
    }
}
