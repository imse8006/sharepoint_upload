﻿# SharePoint XLSX to XLSB Converter & Uploader

## Overview
This script automates the process of converting an Excel file (`.xlsx`) to the binary Excel format (`.xlsb`) and then uploads the converted file to a SharePoint document library. The filename is automatically prefixed with the current date in `YYYYMMDD_` format before uploading.

## Requirements
### 1. Python Libraries
Ensure that the following Python libraries are installed:
- `office365-rest-python-client` (for SharePoint API access)
- `pywin32` (for Excel automation via COM objects)

Install missing dependencies using:
```sh
pip install office365-rest-python-client pywin32
```

### 2. Microsoft Excel (Windows only)
This script requires Microsoft Excel installed on the system, as it leverages the COM interface (`win32com.client`) to perform the file conversion.

## Configuration
### 1. Update SharePoint Credentials
Modify the following variables in the script with your SharePoint details:
```python
SHAREPOINT_SITE_URL = "https://sysco.sharepoint.com/sites/PGMDatabaseSyscoandBain"
CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-client-secret"
TARGET_FOLDER = "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/SPRINT tool/Chris Daniels - SOW/PGM Data"
```

### 2. Set the Input File Path
Specify the `.xlsx` file to be converted and uploaded:
```python
LOCAL_FILE_PATH = "PGM data_Sprint_tool_input.xlsx"
```

## Execution
To run the script, execute the following command:
```sh
python script.py
```

## Workflow
1. **Convert `.xlsx` to `.xlsb`**
   - The script opens the `.xlsx` file using Excel.
   - It saves the file in `.xlsb` format.
   - The new file is prefixed with the current date (e.g., `20240318_PGM data_Sprint_tool_input.xlsb`).

2. **Upload the converted file to SharePoint**
   - The script authenticates with SharePoint using `Client ID` and `Client Secret`.
   - It uploads the `.xlsb` file to the specified SharePoint document library.

## Troubleshooting
### 1. "ModuleNotFoundError: No module named 'win32com'"
Run:
```sh
pip install pywin32
python -m pywin32_postinstall
```

### 2. "403 Forbidden" Error During Upload
- Ensure that the application has sufficient permissions in SharePoint.
- Verify that the `TARGET_FOLDER` path is correctly formatted.

### 3. "File Not Found" Error
- Confirm that the `.xlsx` file exists at the specified path before execution.

## Notes
- The script is **Windows-only** due to the reliance on Excel automation.
- It is recommended to run the script with administrative privileges to avoid permission issues.
- The script does **not** delete the converted file after upload. If needed, uncomment the `os.remove(converted_file_path)` line in the script.

## License
This script is intended for internal use. Modify it as needed for your specific SharePoint environment.

