# CSV to XLSX Converter Windows Service

This application monitors a folder for new CSV files and converts them to Excel XLSX files based on a template and mapping configuration.

## Features

- Monitors a specified folder for new CSV files
- Converts CSV files to Excel (XLSX) format using a predefined template
- Supports custom column and cell mappings
- Can run as both a command line application and a Windows Service
- Logs operations to the Windows Event Log when running as a service

## Configuration

The application uses `appsettings.json` for configuration:

```json
{
  "FilePaths": {
    "TemplatePath": "ExcelTemplate.xlsx",
    "CsvFolderPath": "InputCsv",
    "MappingPath": "mapping.json",
    "OutputFolderPath": "OutputExcel"
  },
  "Encoding": {
    "CsvEncoding": "utf-8"
  }
}
```

- **TemplatePath**: Path to the Excel template file
- **CsvFolderPath**: Path to the folder to monitor for CSV files
- **MappingPath**: Path to the mapping configuration JSON file
- **OutputFolderPath**: Path to save the generated Excel files
- **CsvEncoding**: Character encoding to use when reading CSV files

## Running as a Windows Service

### Installation and Management Options

You have several options to install and manage the Windows Service:

#### Option 1: Using the Interactive Menu (Recommended for Most Users)

1. Right-click on `Manage-Service.ps1` and select "Run with PowerShell"
2. From the menu, select the desired operation (install, start, stop, or uninstall)
3. Accept the UAC prompt to allow administrative privileges

#### Option 2: Using PowerShell Install Script Directly

1. Open PowerShell as Administrator
2. Navigate to the installation directory
3. Run one of the following commands:
   ```powershell
   .\Install-Service.ps1 -Command install
   .\Install-Service.ps1 -Command start
   .\Install-Service.ps1 -Command stop
   .\Install-Service.ps1 -Command uninstall
   ```

#### Option 3: Using ServiceManager.bat (Command Prompt)

1. Open Command Prompt as Administrator
2. Navigate to the installation directory
3. Run one of the following commands:
   ```
   ServiceManager.bat install
   ServiceManager.bat start
   ServiceManager.bat stop
   ServiceManager.bat uninstall
   ```

### Service Operation

Once installed and started, the service will:

1. Monitor the specified input folder for new CSV files
2. Process any existing CSV files that haven't been processed
3. Convert new CSV files to Excel using the specified template and mapping
4. Save the output files to the specified output folder
5. Log all activities to the Windows Event Log

## Running as a Console Application

To run the application in console mode, simply run the executable without any parameters:

```
Csv2Xlsx.exe
```

In console mode, you'll see real-time output in the console window, and the application will run until you press Enter to exit.

## Monitoring and Logging

When running as a Windows Service, all operations are logged to the Windows Event Log under the source name "Csv2XlsxService" in the "Application" log.

To view logs:
1. Open Event Viewer
2. Go to Windows Logs > Application
3. Filter for the source "Csv2XlsxService"

## Troubleshooting

If you encounter issues with the service:

1. Check the Event Log for error messages
2. Ensure the service account has access to all needed files and folders
3. Verify that the configuration in appsettings.json is correct
4. Try running the application in console mode first to check for any issues