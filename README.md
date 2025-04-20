# CSV to XLSX Converter Windows Service

This application monitors a folder for new CSV files and converts them to Excel XLSX files based on a template and mapping configuration.

## Features

- Monitors a specified folder for new CSV files
- Converts CSV files to Excel (XLSX) format using a predefined template
- Supports custom column and cell mappings
- Can run as both a command line application and a Windows Service
- Uses Serilog for comprehensive logging with rolling file appenders

## Configuration

The application uses `appsettings.json` for configuration:

```json
{
  "FilePaths": {
    "TemplatePath": "ExcelTemplate.xlsx",
    "CsvFolderPath": "CsvFiles",
    "MappingPath": "mapping.json",
    "OutputFolderPath": "ProcessedFiles"
  },
  "Encoding": {
    "CsvEncoding": "windows-1255"
  },
  "Serilog": {
    "MinimumLevel": {
      "Default": "Information",
      "Override": {
        "Microsoft": "Warning",
        "System": "Warning"
      }
    }
  }
}
```

- **TemplatePath**: Path to the Excel template file
- **CsvFolderPath**: Path to the folder to monitor for CSV files
- **MappingPath**: Path to the mapping configuration JSON file
- **OutputFolderPath**: Path to save the generated Excel files
- **CsvEncoding**: Character encoding to use when reading CSV files (e.g., windows-1255, utf-8)
- **Serilog**: Configuration for logging

## Running as a Windows Service

### Installation and Management

The simplest way to install and manage the Windows Service is using the included batch script:

1. Right-click on `ServiceManager-Direct.bat` and select "Run as administrator"
2. From the menu, select one of the following options:
   - **Install Service** - Registers the application as a Windows Service
   - **Start Service** - Starts the service
   - **Stop Service** - Stops the service
   - **Uninstall Service** - Removes the service
   - **Check Service Status** - Shows the current status of the service
   - **View Log Files** - Opens the logs directory

### Service Operation

Once installed and started, the service will:

1. Monitor the specified input folder (CsvFiles) for new CSV files
2. Process any existing CSV files that haven't been processed
3. Convert new CSV files to Excel using the specified template and mapping
4. Save the output files to the specified output folder (ProcessedFiles)
5. Log all activities to the logs directory with automatic file rotation

## Running as a Console Application

To run the application in console mode, simply run the executable without any parameters:

```
Csv2Xlsx.exe
```

In console mode, you'll see real-time output in the console window, and the application will run until you press Enter to exit.

## Logging

The application uses Serilog for structured logging with the following features:

- Logs are stored in the `logs` directory within the application folder
- Log files are named with the pattern `csv2xlsx-YYYYMMDD.log`
- Files rotate daily and when they reach 10MB in size
- Logs are kept for 31 days before being automatically deleted
- Detailed, structured log entries with timestamps and log levels

When running in console mode, logs are displayed in the console window in addition to being written to log files.

To view logs:
1. Use the "View Log Files" option in the ServiceManager-Direct.bat menu, or
2. Navigate to the logs directory in the application folder

## Mapping Configuration

The mapping between CSV and Excel is defined in `mapping.json`. The format is:

```json
{
  "sheetName": "Sheet1",
  "startRow": 2,
  "columnMappings": {
    "CSVColumn1": "A",
    "CSVColumn2": "B"
  },
  "cellMappings": [
    {
      "csvColumn": "TotalValue",
      "csvRow": 1,
      "excelColumn": "F",
      "excelRow": 15
    }
  ]
}
```

- **sheetName**: The Excel worksheet to modify
- **startRow**: The row number to start writing data (1-based)
- **columnMappings**: Maps CSV columns to Excel columns
- **cellMappings**: Maps specific CSV values to specific Excel cells

## Troubleshooting

If you encounter issues with the service:

1. Check the log files for error messages
2. Ensure the service account (LocalSystem) has access to all needed files and folders
3. Verify that the configuration in appsettings.json is correct
4. Try running the application in console mode first to check for any issues
5. Use the "Check Service Status" option to verify the service is running