# Technical Documentation

This document contains detailed technical information about the Webcon PowerShell Workflow Automation tool.

## Table of Contents

- [Configuration Details](#configuration-details)
- [Excel File Format](#excel-file-format)
- [Field Type Detection](#field-type-detection)
- [Status Tracking](#status-tracking)
- [Retry Logic](#retry-logic)
- [Environment Variables](#environment-variables)
- [Module Functions](#module-functions)
- [Error Handling](#error-handling)
- [Advanced Customization](#advanced-customization)
- [Troubleshooting](#troubleshooting)

## Configuration Details

### Config.json Structure

```json
{
  "Webcon": {
    "BaseUrl": "https://your-webcon-instance.com",
    "ClientId": "your-client-id-here",
    "ClientSecret": "",
    "DatabaseId": "9"
  },
  "Workflow": {
    "WorkflowGuid": "your-workflow-guid-here",
    "FormTypeGuid": "your-form-type-guid-here",
    "Path": "default",
    "Mode": "standard"
  },
  "Excel": {
    "FilePath": "C:\\data\\workflows.xlsx",
    "StartRow": 6
  },
  "StatusFile": "",
  "Retry": {
    "MaxRetries": 3,
    "RetryDelayBase": 2
  },
  "ItemList": {
    "Enabled": false,
    "SheetName": "ItemList",
    "ItemListGuid": "your-item-list-guid-here",
    "ItemListName": "your-item-list-name-here"
  }
}
```

### Webcon Settings
- `BaseUrl`: Your Webcon instance URL
- `ClientId`: OAuth2 client ID
- `ClientSecret`: OAuth2 client secret (optional - can be left empty if using environment variable)
- `DatabaseId`: Database ID number

**Security Note**: For security, it's recommended to store the `ClientSecret` in an environment variable instead of the config file. See [Environment Variables](#environment-variables) section below.

### Workflow Settings
- `WorkflowGuid`: The GUID of the workflow to start
- `FormTypeGuid`: The GUID of the form type
- `Path`: Path parameter (default: "default" if not provided)
- `Mode`: Mode parameter (default: "standard" if not provided)

### Excel Settings
- `FilePath`: Full path to your Excel file
- `StartRow`: Row number to start reading data (default: 5, rows 1-4 are metadata)

### ItemList Settings (Optional)
- `Enabled`: Set to `true` to enable item list import (default: `false`)
- `SheetName`: Name of the item list sheet (default: "ItemList")
- `ItemListGuid`: The GUID of the item list (required if Enabled is true)
- `ItemListName`: The name of the item list (required if Enabled is true)

### Status Tracking (Optional)
- `StatusFile`: Path to status CSV file (default: same directory as Excel file with `.status.csv` extension)
  - If empty, automatically creates: `{ExcelFileName}.status.csv` in the same directory as the Excel file

### Retry Settings (Optional)
- `Retry.MaxRetries`: Maximum number of retry attempts for transient errors (default: 3)
- `Retry.RetryDelayBase`: Base delay in seconds for exponential backoff (default: 2)
  - Retry delays: 2s, 4s, 8s, etc. (exponential backoff)

## Excel File Format

Your Excel file **must have one sheet** named "Data" (plus optional "ItemList" sheet if importing item lists).

### Sheet: "Data"
This sheet contains both field mappings (rows 1-4) and data rows (row 5+).

**Rows 1-4: Field Metadata** (populated from SQL stored procedure results)
These rows define the field mappings between Excel columns and Webcon form fields.

- **Row 1**: Friendly field names (e.g., "Active", "Company Name", "Email") - *From SQL column headers*
- **Row 2**: Database/Technical names (e.g., "WFD_AttBool1", "WFD_AttText1", "WFD_AttText2") - *From SQL stored procedure row 1*
- **Row 3**: Field GUIDs (the GUID of each Webcon form field) - *From SQL stored procedure row 2*
- **Row 4**: ColumnType (e.g., "Yes / No choice", "Single line of text", "Floating-point number", "Multiple lines of text") - *From SQL stored procedure row 3*

**Row 5+: Data Rows**
Your actual data starts from row 5. Column names in row 5+ should match the column headers from row 1.

**Structure:**
- **Row 1**: Friendly names (e.g., "Active", "CompanyName", "Email") - *From SQL column headers*
- **Row 2**: Database names (e.g., "WFD_AttBool1", "WFD_AttText1", "WFD_AttText2") - *From SQL stored procedure row 1*
- **Row 3**: GUIDs (e.g., "89e652c8-f338-49f3-bc84-24f622a2eb88", ...) - *From SQL stored procedure row 2*
- **Row 4**: ColumnType (e.g., "Yes / No choice", "Single line of text", "Floating-point number") - *From SQL stored procedure row 3*
- **Row 5+**: Data rows

**Optional columns:**
- `ID`: Column to uniquely identify rows. Used for status tracking and linking item list rows to workflow instances. If not present, row numbers are used.

**Example Data sheet:**
| Active | CompanyName | Email |
|--------|-------------|-------|
| WFD_AttBool1 | WFD_AttText1 | WFD_AttText2 |
| 89e652c8-f338-49f3-bc84-24f622a2eb88 | 2c7354b6-7b90-4a32-9471-6f39d68d5458 | 6c610728-24db-4a5d-a704-b33f849a1f05 |
| Yes / No choice | Single line of text | Single line of text |
| ID | Active | CompanyName | Email |
| 1 | Yes | ACME d.o.o. | info@acme.com |
| 2 | No | Another Company | contact@another.com |

### Optional Sheet: "ItemList"
If you're importing item lists, create a sheet with the same structure as the Data sheet.

**Rows 1-4: Item List Column Metadata**
- **Row 1**: Friendly column names (e.g., "Account Number", "Discount Value") - *From SQL column headers*
- **Row 2**: Database/Technical names (e.g., "DET_Att1", "DET_Value1") - *From SQL stored procedure row 1*
- **Row 3**: Column GUIDs - *From SQL stored procedure row 2*
- **Row 4**: ColumnType (e.g., "Single line of text", "Floating-point number") - *From SQL stored procedure row 3*

**Row 5+: Item List Data Rows**
- Must include an `ID` column that matches the `ID` from the Data sheet to link item list rows to workflow instances.

**Example ItemList sheet:**
| RowType | Account Number | Discount Value |
|---------|----------------|----------------|
| DatabaseName | DET_Att1 | DET_Value1 |
| Guid | BA0F6ADB-CDF6-4066-A287-B977A354EA1B | B0DF394B-2F36-4A5F-ABE7-BB93A6266F85 |
| ColumnType | Single line of text | Floating-point number |
| ID | Account Number | Discount Value |
| 1 | ACC001 | 10.5 |
| 1 | ACC002 | 15.0 |
| 2 | ACC003 | 20.0 |

## Field Type Detection

The script automatically detects field types using a two-tier approach:

**Primary Method: ColumnType Parsing**
The script detects field types - Primary: ColumnType parsing, Secondary: DatabaseName patterns:
- "Yes / No choice" → Boolean field
- "Floating-point number" → Decimal field
- "Single line of text" → String field
- "Multiple lines of text" → String/LongText field
- Contains "choice" (case-insensitive, excluding "Yes / No choice") → Choice field

**Secondary Method: DatabaseName Pattern Matching**
If ColumnType doesn't match known patterns, the script falls back to DatabaseName patterns:

### Choice Fields
**Detection**: ColumnType contains "choice" (excluding "Yes / No choice") OR DatabaseName contains "Choose" or "Choice" (e.g., `WFD_AttChoose2`, `DET_AttChoose1`)
- **Value format**: 
  - Single value (used as id, name left blank): `0000019`
  - Or: `id#name` format: `0000019#Customer Name`
- **API format**: Array of objects with `id` and `name` properties

### Boolean Fields
**Detection**: ColumnType is "Yes / No choice" OR DatabaseName contains "AttBool" (e.g., `WFD_AttBool1`, `DET_AttBool1`)
- **Value format**: `true`, `false`, `1`, `0`, `yes`, `no`, `y`, `n`
- **Converted to**: Boolean `true` or `false`
- **API format**: Boolean value (not string)

### DateTime Fields
**Detection**: DatabaseName contains "AttDateTime" (e.g., `WFD_AttDateTime2`, `DET_AttDateTime1`)
- **Value format**: Any valid DateTime format (Excel date, ISO string, etc.)
- **Converted to**: ISO 8601 format: `2025-11-05T12:42:24.305Z`
- **API format**: ISO 8601 string

### Integer Fields
**Detection**: DatabaseName contains "AttInt" (e.g., `WFD_AttInt1`, `DET_AttInt1`)
- **Value format**: Integer number
- **Converted to**: Integer: `0`, `1`, `123`, etc.
- **API format**: Integer number

### Decimal Fields
**Detection**: ColumnType is "Floating-point number" OR DatabaseName contains "AttDecimal" or starts with "DET_Value" (e.g., `WFD_AttDecimal7`, `DET_Value1`)
- **Value format**: Decimal number
- **Converted to**: Decimal: `0`, `1.5`, `123.45`, etc.
- **API format**: Decimal number

### Long Text Fields
**Detection**: ColumnType is "Multiple lines of text" OR DatabaseName starts with "DET_LongText" (item lists only)
- **Value format**: Any text (can contain newlines)
- **API format**: String value

### Regular Fields
All other fields are treated as strings
- **Value format**: Any text
- **API format**: String value

## Status Tracking

The script maintains a separate CSV status file that tracks the import status of each row:

### Status File Structure
- **Location**: Same directory as Excel file, named `{ExcelFileName}.status.csv`
- **Format**: CSV with columns: `ID`, `Status`, `ImportedDate`, `ErrorMessage`
- **Status Values**: 
  - `Success`: Row successfully imported
  - `Error`: Row failed to import (with error message)
  - `NotStarted`: Row not yet processed
  - `Metadata`: Special rows for start/end timestamps (`__START__` and `__END__`)

### Status File Format
```
ID,Status,ImportedDate,ErrorMessage
__START__,Metadata,2024-01-15 10:30:00,
1,Success,2024-01-15 10:31:00,
2,Error,2024-01-15 10:32:00,Some error message
3,Success,2024-01-15 10:33:00,
__END__,Metadata,2024-01-15 10:35:00,
```

### Resume Capability
- **Automatic Resume**: On rerun, the script automatically skips rows with `Status = Success`
- **Retry Failed Rows**: Rows with `Status = Error` will be retried on the next run
- **Progress Preservation**: If the script stops (timeout, crash, etc.), you can simply rerun it to continue from where it left off
- **Start/End Timestamps**: The status file includes `__START__` and `__END__` rows with execution timestamps

### Resetting Status
To restart an import from scratch:
1. Delete or rename the `.status.csv` file
2. Or manually edit the CSV to change status values back to `NotStarted`

## Retry Logic

The script automatically retries transient errors with exponential backoff:

### Retryable Errors
- **Network timeouts**: Connection timeouts, request timeouts
- **Network failures**: Connection failures, receive failures
- **Server errors**: HTTP 500, 502, 503, 504 (server-side issues)
- **Timeout exceptions**: General timeout exceptions

### Non-Retryable Errors (Permanent Failures)
- **Client errors**: HTTP 400, 401, 403, 404 (validation, authentication, authorization errors)
- These are logged immediately without retry

### Retry Behavior
- **Default**: 3 retry attempts with exponential backoff
- **Delays**: 2 seconds, 4 seconds, 8 seconds (configurable)
- **Formula**: `delay = RetryDelayBase * (2 ^ attemptNumber)`
- **Example**: With MaxRetries=3 and RetryDelayBase=2:
  - Attempt 1: Immediate
  - Attempt 2: Wait 2 seconds
  - Attempt 3: Wait 4 seconds
  - Attempt 4: Wait 8 seconds
  - If all fail: Mark as Error in status file

## Environment Variables

For better security, you can store the `ClientSecret` in an environment variable instead of the `Config.json` file.

### Setting the Environment Variable

#### Windows (PowerShell - Current Session Only)
```powershell
$env:WEBCON_CLIENT_SECRET = "your-client-secret-here"
```

#### Windows (PowerShell - Permanent for Current User)
```powershell
[System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-client-secret-here", "User")
```

**Note**: The script automatically reads from the User registry if the process environment variable is not set, so you don't need to restart PowerShell after setting it with `SetEnvironmentVariable`.

#### Windows (Command Prompt - Current Session Only)
```cmd
set WEBCON_CLIENT_SECRET=your-client-secret-here
```

#### Windows (Command Prompt - Permanent for Current User)
```cmd
setx WEBCON_CLIENT_SECRET "your-client-secret-here"
```

#### Windows (System-Wide - Requires Admin)
```powershell
[System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-client-secret-here", "Machine")
```

**Note**: After setting a permanent environment variable with `setx` or for Machine scope, you may need to:
- Close and reopen PowerShell/Command Prompt
- Or restart your computer (for system-wide variables)

### How It Works

1. The script first checks if `ClientSecret` is provided in `Config.json`
2. If empty or not provided, it reads from the `WEBCON_CLIENT_SECRET` process environment variable
3. If not found in process environment, it reads from the User registry (where `SetEnvironmentVariable` stores permanent variables)
4. If neither is found, the script will throw an error

### Example Config.json

```json
{
  "Webcon": {
    "BaseUrl": "https://your-webcon-instance.com",
    "ClientId": "your-client-id-here",
    "ClientSecret": "",
    "DatabaseId": "9"
  },
  "Workflow": {
    "WorkflowGuid": "your-workflow-guid-here",
    "FormTypeGuid": "your-form-type-guid-here",
    "Path": "default",
    "Mode": "standard"
  },
  "Excel": {
    "FilePath": "C:\\data\\workflows.xlsx",
    "StartRow": 6
  },
  "ItemList": {
    "Enabled": false,
    "SheetName": "ItemList",
    "ItemListGuid": "your-item-list-guid-here",
    "ItemListName": "your-item-list-name-here"
  }
}
```

Leave `ClientSecret` empty in the config file when using environment variables.

## Module Functions

### WebconAPI.psm1
- `Get-WebconToken`: Authenticates and returns access token
- `Start-WebconWorkflow`: Creates a new workflow element
- `Start-WebconWorkflowWithRetry`: Creates a workflow element with automatic retry logic for transient errors

### ExcelReader.psm1
- `Read-FieldMappingsFromDataSheet`: Reads field mappings from rows 1-4 of the Data sheet
- `Read-ExcelFile`: Reads data rows from the Data sheet (starting at row 5)
- `Read-ItemListMappingsFromDataSheet`: Reads item list column mappings from rows 1-4 of the ItemList sheet
- `Read-ItemListData`: Reads item list data rows from the ItemList sheet (starting at row 5)

### StatusTracker.psm1
- `Get-ImportStatus`: Reads status CSV file and returns hashtable of row IDs and their status
- `Update-ImportStatus`: Writes status to CSV file for a specific row ID
- `IsRowImported`: Checks if a row ID has already been successfully imported
- `Write-StartMetadata`: Writes start timestamp to status CSV file
- `Write-EndMetadata`: Writes end timestamp to status CSV file

### ProgressWindow.psm1
- `Show-ProgressWindow`: Creates and displays a Windows Forms progress window with real-time updates

## Error Handling

The script continues processing even if individual rows fail. Errors are:
- Displayed in red during execution
- Collected and summarized at the end
- Do not stop the entire process
- Logged to the status CSV file with error messages

## Advanced Customization

### Adding New Workflows

For different workflows, you can:
- Update the `Workflow` section in `Config.json` with different workflow and form type GUIDs
- Update rows 1-4 of the Data sheet with the appropriate field GUIDs and metadata (use SQL stored procedures to get mappings)
- Update the Data sheet columns to match your field mappings
- Or create a new Excel file with the Data sheet configured

### Changing Workflow Configuration

Update the `Workflow` section in `Config.json`:
- Change `WorkflowGuid` to target a different workflow
- Change `FormTypeGuid` to use a different form type
- Modify `Path` or `Mode` if needed

### Changing Field Mappings

Update rows 1-4 of the Data sheet (use SQL stored procedures to get mappings):
- **Row 1**: Update friendly names (add/remove columns) - *From SQL column headers*
- **Row 2**: Update database/technical names - *From SQL stored procedure row 1*
- **Row 3**: Update field GUIDs - *From SQL stored procedure row 2*
- **Row 4**: Update ColumnType values - *From SQL stored procedure row 3*
- Add new columns to map additional Excel columns to Webcon fields
- Remove columns to exclude fields

## Troubleshooting

### "ImportExcel module not found"
Run: `Install-Module ImportExcel -Scope CurrentUser`

### "Failed to get access token"
- Verify your `ClientId` and `ClientSecret` are correct
- Check that the `BaseUrl` is correct
- Ensure your OAuth2 application has the necessary permissions

### "ClientSecret not found"
- If using environment variable: `[System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-secret", "User")`
- The script will automatically read from User registry, no restart needed
- Or provide it in `Config.json` (less secure)

### "Failed to start workflow"
- Verify workflow and form type GUIDs are correct
- Check that all required fields are mapped
- **Make sure the system user (OAuth2 client) used by the script has start workflow privileges in Webcon**
- Review the error response body for details
- Check the progress window for specific error messages

### "Excel file not found"
- Verify the `FilePath` in Config.json is correct
- Use absolute paths if relative paths don't work
- Ensure the file is not locked by another application

### "Workflow configuration not found in Config.json"
- Ensure the `Workflow` section exists in `Config.json`
- Check that `WorkflowGuid` and `FormTypeGuid` are provided
- Verify the JSON syntax is correct

### "Data sheet is missing metadata rows"
- Ensure rows 1-4 exist in the Data sheet
- Check that row 1 contains friendly field names
- Verify that row 2 contains database/technical names
- Verify that row 3 contains field GUIDs
- Verify that row 4 contains ColumnType values

### "Data sheet not found"
- Ensure your Excel file has a sheet named "Data"
- Check that the sheet name is spelled exactly as "Data" (case-sensitive)

### Import stops mid-process
- **Resume**: Simply rerun the script - it will automatically skip successfully imported rows
- **Check status file**: Review the `.status.csv` file to see which rows succeeded/failed
- **Retry failed rows**: Delete the status file or change Error status to NotStarted to retry

### Timeouts or network errors
- The script automatically retries transient errors (timeouts, network issues)
- Check the retry settings in Config.json if you need to adjust retry behavior
- Permanent errors (authentication, validation) are not retried - fix the underlying issue

### Progress window issues
- If the progress window doesn't appear, ensure Windows Forms is available
- The window should stay on top and update in real-time
- Close the window when processing completes to finish the script

### JSON encoding errors
- The script handles UTF-8 encoding automatically
- Special characters (like ø, é, etc.) should work correctly
- If you see encoding errors, check that your Excel file is saved with proper encoding

## Security Notes

- **Recommended**: Store `ClientSecret` in environment variable (`WEBCON_CLIENT_SECRET`) instead of `Config.json`
- Store `Config.json` securely or use environment variables for sensitive data
- Never commit credentials to version control
- Environment variables are more secure as they're not stored in files
- The status CSV file may contain error messages - review before sharing

## How It Works

1. **Load Status**: Loads import status from CSV file (if exists) to track already-imported rows
2. **Read Workflow Config**: Reads workflow configuration from `Config.json` (Workflow section)
3. **Read Field Mappings**: Reads field mappings from rows 1-4 of the Data sheet
4. **Read Item List Mappings** (if enabled): Reads item list column mappings from rows 1-4 of the ItemList sheet
5. **Authentication**: Authenticates with Webcon using OAuth2 client credentials
6. **Read Data**: Reads data rows from the Data sheet (starting at row 5)
7. **Read Item List Data** (if enabled): Reads item list data rows from the ItemList sheet (starting at row 5) and groups by ID
8. **Show Progress Window**: Displays a Windows Forms window with real-time progress
9. **Row Processing**: For each data row:
   - Checks if row is already successfully imported (skips if yes)
   - Maps Excel columns to Webcon form fields using the mappings from rows 1-4
   - Finds matching item list rows (if enabled) using ID
   - Creates one workflow element via Webcon API with automatic retry for transient errors
   - Updates status CSV file with result (Success or Error)
   - Updates progress window with current status
10. **Summary**: Displays success/error/skipped counts and details
11. **Wait for User**: Progress window remains open for review until user closes it

## License

This is a custom solution for internal use.

