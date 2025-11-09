# Webcon PowerShell Workflow Automation

This PowerShell solution reads Excel files and automatically starts workflows in Webcon for each row. It's designed to be modular and reusable for different workflows and field mappings.

## Features

- **Modular Design**: Separate modules for Excel reading and Webcon API operations
- **Excel-Based Mappings**: Field mappings stored in Excel for easy configuration
- **Configurable**: JSON-based configuration for connection and workflow settings
- **Resume Capability**: Automatically skips already-imported rows, resume from where you left off
- **Status Tracking**: Separate CSV file tracks import status for each row
- **Retry Logic**: Automatic retry with exponential backoff for transient errors (timeouts, network issues)
- **Error Handling**: Continues processing even if individual rows fail
- **One API Call Per Row**: Simple, efficient processing with one workflow creation per data row
- **Detailed Logging**: Color-coded output for easy monitoring

## Prerequisites

- PowerShell 5.1 or later
- ImportExcel PowerShell module (automatically installed if missing)
- Access to Webcon API with OAuth2 credentials

## Installation

1. Install the required PowerShell module:
   ```powershell
   Install-Module ImportExcel -Scope CurrentUser
   ```

2. Ensure all files are in place:
   ```
   webcon-powershell/
   ├── Modules/
   │   ├── WebconAPI.psm1
   │   ├── ExcelReader.psm1
   │   └── StatusTracker.psm1
   ├── Config.json
   ├── Start-WebconWorkflows.ps1
   └── README.md
   ```

## Configuration

Edit `Config.json` to configure:

### Webcon Settings
- `BaseUrl`: Your Webcon instance URL
- `ClientId`: OAuth2 client ID
- `ClientSecret`: OAuth2 client secret (optional - can be left empty if using environment variable)
- `DatabaseId`: Database ID number

**Security Note**: For security, it's recommended to store the `ClientSecret` in an environment variable instead of the config file. See [Environment Variables](#environment-variables) section below.

### Workflow Settings
**Note**: Workflow settings are now stored in the Excel file's "Mapping-Workflow" sheet. See [Excel File Format](#excel-file-format) below.

### Excel Settings
- `FilePath`: Full path to your Excel file
- `StartRow`: Row number to start reading (default: 2, assumes row 1 is headers)

### Status Tracking (Optional)
- `StatusFile`: Path to status CSV file (default: same directory as Excel file with `.status.csv` extension)
  - If empty, automatically creates: `{ExcelFileName}.status.csv` in the same directory as the Excel file

### Retry Settings (Optional)
- `Retry.MaxRetries`: Maximum number of retry attempts for transient errors (default: 3)
- `Retry.RetryDelayBase`: Base delay in seconds for exponential backoff (default: 2)
  - Retry delays: 2s, 4s, 8s, etc. (exponential backoff)

## Excel File Format

Your Excel file **must have three sheets** with fixed names:

### Sheet 1: "Mapping-Workflow"
This sheet defines the workflow configuration.

**Required columns:**
- `WorkflowGuid`: The GUID of the workflow to start
- `FormTypeGuid`: The GUID of the form type

**Optional columns:**
- `Path`: Path parameter (default: "default" if not provided)
- `Mode`: Mode parameter (default: "standard" if not provided)

**Structure:**
- **Row 1**: Headers (WorkflowGuid, FormTypeGuid, Path, Mode)
- **Row 2**: Configuration data (only first row is used)

**Example Mapping-Workflow sheet:**
| WorkflowGuid | FormTypeGuid | Path | Mode |
|--------------|--------------|------|------|
| f395d755-5a7b-4624-8169-869e5a149b5b | 810ac36c-f605-4762-8ccd-52ec42288c77 | default | standard |

### Sheet 2: "Mapping-Fields"
This sheet defines the field mappings between Excel columns and Webcon form fields.

**Required columns:**
- `ExcelColumn`: The column name from the Data sheet
- `FieldGuid`: The GUID of the Webcon form field
- `FieldName`: The name of the Webcon form field
- `FieldType`: The type of the Webcon form field (e.g., "Unspecified")

**Optional columns:**
- `IsChoice`: Set to "Yes", "True", or "1" for choice/dropdown fields (optional - auto-detected if FieldName contains "Choose" or "Choice")

**Structure:**
- **Row 1**: Headers (ExcelColumn, FieldGuid, FieldName, FieldType, IsChoice)
- **Row 2+**: Mapping data

**Example Mapping-Fields sheet:**
| ExcelColumn | FieldGuid | FieldName | FieldType | IsChoice |
|-------------|-----------|-----------|-----------|----------|
| CompanyName | 3712b43b-5947-4c7b-b73a-372ea83daa91 | WFD_AttText1 | Unspecified | |
| Customer | 331bfbca-0bc2-47f6-8745-02ae38895e8f | WFD_AttChoose2 | Unspecified | Yes |

**Field Type Detection:**
The script automatically detects field types based on FieldName patterns:

- **Choice Fields**: Fields with "Choose" or "Choice" in the name (e.g., `WFD_AttChoose2`)
  - Value format: Single value (used as id, name left blank): `0000019`
  - Or: `id#name` format: `0000019#Customer Name`
  - Can also be marked with `IsChoice` column set to "Yes"

- **Boolean Fields**: Fields with "AttBool" in the name (e.g., `WFD_AttBool1`)
  - Value format: `true`, `false`, `1`, `0`, `yes`, `no`, `y`, `n`
  - Converted to boolean: `true` or `false`

- **DateTime Fields**: Fields with "AttDateTime" in the name (e.g., `WFD_AttDateTime2`)
  - Value format: Any valid DateTime format (Excel date, ISO string, etc.)
  - Converted to ISO 8601 format: `2025-11-05T12:42:24.305Z`

- **Integer Fields**: Fields with "AttInt" in the name (e.g., `WFD_AttInt1`)
  - Value format: Integer number
  - Converted to integer: `0`, `1`, `123`, etc.

- **Decimal Fields**: Fields with "AttDecimal" in the name (e.g., `WFD_AttDecimal7`)
  - Value format: Decimal number
  - Converted to decimal: `0`, `1.5`, `123.45`, etc.

- **Regular Fields**: All other fields are treated as strings

### Sheet 3: "Data"
This sheet contains the actual data rows to process.

**Structure:**
- **Row 1**: Column headers (must match `ExcelColumn` values from the Mapping-Fields sheet)
- **Row 2+**: Data rows
- **Optional**: Include an `ID` column to uniquely identify rows. If not present, row numbers are used.

**Example Data sheet:**
| ID | CompanyName | Email |
|----|-------------|-------|
| 1  | ACME d.o.o. | info@acme.com |
| 2  | Another Company | contact@another.com |

## Usage

Run the main script:
```powershell
.\Start-WebconWorkflows.ps1
```

Or specify a custom config file:
```powershell
.\Start-WebconWorkflows.ps1 -ConfigPath ".\CustomConfig.json"
```

## How It Works

1. **Load Status**: Loads import status from CSV file (if exists) to track already-imported rows
2. **Read Workflow Config**: Reads workflow configuration from the "Mapping-Workflow" sheet (first sheet)
3. **Read Field Mappings**: Reads field mappings from the "Mapping-Fields" sheet (second sheet)
4. **Authentication**: Authenticates with Webcon using OAuth2 client credentials
5. **Read Data**: Reads data rows from the "Data" sheet (third sheet)
6. **Row Processing**: For each data row:
   - Checks if row is already successfully imported (skips if yes)
   - Maps Excel columns to Webcon form fields using the mappings from the Mapping-Fields sheet
   - Creates one workflow element via Webcon API with automatic retry for transient errors
   - Updates status CSV file with result (Success or Error)
7. **Summary**: Displays success/error/skipped counts and details

## Status Tracking and Resume

The script maintains a separate CSV status file that tracks the import status of each row:

### Status File Structure
- **Location**: Same directory as Excel file, named `{ExcelFileName}.status.csv`
- **Format**: CSV with columns: `ID`, `Status`, `ImportedDate`, `ErrorMessage`
- **Status Values**: 
  - `Success`: Row successfully imported
  - `Error`: Row failed to import (with error message)
  - `NotStarted`: Row not yet processed

### Resume Capability
- **Automatic Resume**: On rerun, the script automatically skips rows with `Status = Success`
- **Retry Failed Rows**: Rows with `Status = Error` will be retried on the next run
- **Progress Preservation**: If the script stops (timeout, crash, etc.), you can simply rerun it to continue from where it left off

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

## Error Handling

The script continues processing even if individual rows fail. Errors are:
- Displayed in red during execution
- Collected and summarized at the end
- Do not stop the entire process

## Customization

### Adding New Workflows

For different workflows, you can:
- Update the "Mapping-Workflow" sheet in your Excel file with different workflow and form type GUIDs
- Update the "Mapping-Fields" sheet with the appropriate field GUIDs
- Update the Data sheet with the columns matching your Mapping-Fields sheet
- Or create a new Excel file with all three sheets configured

### Changing Workflow Configuration

Update the "Mapping-Workflow" sheet in your Excel file:
- Change `WorkflowGuid` to target a different workflow
- Change `FormTypeGuid` to use a different form type
- Modify `Path` or `Mode` if needed

### Changing Field Mappings

Simply update the "Mapping-Fields" sheet in your Excel file:
- Add new rows to map additional Excel columns to Webcon fields
- Remove rows to exclude fields
- Update GUIDs to change which Webcon fields receive the data

## Troubleshooting

### "ImportExcel module not found"
Run: `Install-Module ImportExcel -Scope CurrentUser`

### "Failed to get access token"
- Verify your `ClientId` and `ClientSecret` are correct
- Check that the `BaseUrl` is correct
- Ensure your OAuth2 application has the necessary permissions

### "Failed to start workflow"
- Verify workflow and form type GUIDs are correct
- Check that all required fields are mapped
- Review the error response body for details

### "Excel file not found"
- Verify the `FilePath` in Config.json is correct
- Use absolute paths if relative paths don't work
- Ensure the file is not locked by another application

### "Mapping-Workflow sheet is missing required columns"
- Ensure the Mapping-Workflow sheet (first sheet) has these exact column names: WorkflowGuid, FormTypeGuid
- Check that Row 1 contains the headers
- Ensure Row 2 contains the workflow configuration values

### "Mapping-Fields sheet is missing required columns"
- Ensure the Mapping-Fields sheet (second sheet) has these exact column names: ExcelColumn, FieldGuid, FieldName, FieldType
- Check that Row 1 contains the headers

### "Data sheet not found"
- Ensure your Excel file has a sheet named "Data" (third sheet)
- Check that the sheet name is spelled exactly as "Data"

### Import stops mid-process
- **Resume**: Simply rerun the script - it will automatically skip successfully imported rows
- **Check status file**: Review the `.status.csv` file to see which rows succeeded/failed
- **Retry failed rows**: Delete the status file or change Error status to NotStarted to retry

### Timeouts or network errors
- The script automatically retries transient errors (timeouts, network issues)
- Check the retry settings in Config.json if you need to adjust retry behavior
- Permanent errors (authentication, validation) are not retried - fix the underlying issue

## Module Functions

### WebconAPI.psm1
- `Get-WebconToken`: Authenticates and returns access token
- `Start-WebconWorkflow`: Creates a new workflow element
- `Start-WebconWorkflowWithRetry`: Creates a workflow element with automatic retry logic for transient errors

### ExcelReader.psm1
- `Read-WorkflowMapping`: Reads the Mapping-Workflow sheet (first sheet) and returns workflow configuration
- `Read-MappingSheet`: Reads the Mapping-Fields sheet (second sheet) and returns field mappings
- `Read-ExcelFile`: Reads the Data sheet and returns rows as objects

### StatusTracker.psm1
- `Get-ImportStatus`: Reads status CSV file and returns hashtable of row IDs and their status
- `Update-ImportStatus`: Writes status to CSV file for a specific row ID
- `IsRowImported`: Checks if a row ID has already been successfully imported

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

**Note**: After setting a permanent environment variable, you may need to:
- Close and reopen PowerShell/Command Prompt
- Or restart your computer (for system-wide variables)

### How It Works

1. The script first checks if `ClientSecret` is provided in `Config.json`
2. If empty or not provided, it reads from the `WEBCON_CLIENT_SECRET` environment variable
3. If neither is found, the script will throw an error

### Example Config.json

```json
{
  "Webcon": {
    "BaseUrl": "https://your-webcon-instance.com",
    "ClientId": "your-client-id-here",
    "ClientSecret": "",
    "DatabaseId": "9"
  }
}
```

Leave `ClientSecret` empty in the config file when using environment variables.

## Security Notes

- **Recommended**: Store `ClientSecret` in environment variable (`WEBCON_CLIENT_SECRET`) instead of `Config.json`
- Store `Config.json` securely or use environment variables for sensitive data
- Never commit credentials to version control
- Environment variables are more secure as they're not stored in files

## License

This is a custom solution for internal use.

