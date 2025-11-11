# Webcon PowerShell Workflow Automation

Automatically import data from Excel files into Webcon workflows. This tool reads Excel files and creates workflow elements in Webcon for each row.

## Quick Start

### 1. Prerequisites
- PowerShell 5.1 or later
- ImportExcel module (installed automatically if missing)

### 2. Installation
   ```powershell
   Install-Module ImportExcel -Scope CurrentUser
   ```

### 3. Configuration

Edit `Config.json`:
```json
{
  "Webcon": {
    "BaseUrl": "https://your-webcon-instance.com",
    "ClientId": "your-client-id-here",
    "ClientSecret": "",
    "DatabaseId": "9"
  },
  "Excel": {
    "FilePath": "C:\\data\\workflows.xlsx",
    "StartRow": 2
  }
}
```

**Security**: Leave `ClientSecret` empty and set it as an environment variable:
```powershell
[System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-secret", "User")
```

**Important**: Make sure the system user (OAuth2 client) used by the script to call the API has **start workflow privileges** in Webcon.

### 4. Prepare Excel File

Your Excel file needs **3 sheets**:

**Sheet 1: "Mapping-Workflow"** - Workflow configuration
| WorkflowGuid | FormTypeGuid | Path | Mode |
|--------------|--------------|------|------|
| your-guid-here | your-guid-here | default | standard |

**Sheet 2: "Mapping-Fields"** - Field mappings
| ExcelColumn | FieldGuid | FieldName | FieldType |
|-------------|-----------|-----------|-----------|
| CompanyName | guid-here | WFD_AttText1 | Unspecified |

**Sheet 3: "Data"** - Your data rows
| ID | CompanyName | Email |
|----|-------------|-------|
| 1  | ACME Inc.   | info@acme.com |

### 5. Run
```powershell
.\Start-WebconWorkflows.ps1
```

A progress window will show real-time status. The script automatically:
- Skips already imported rows (resume capability)
- Retries on transient errors
- Tracks status in a `.status.csv` file

## Basic Troubleshooting

**"ClientSecret not found"**
- Set environment variable: `[System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-secret", "User")`
- Or provide it in `Config.json` (less secure)

**"Excel file not found"**
- Use absolute path in `Config.json`
- Ensure file is not locked by Excel

**"Failed to start workflow"**
- Verify workflow and form type GUIDs in Mapping-Workflow sheet
- Check that all required fields are mapped

**Import stopped mid-process?**
- Just rerun the script - it automatically resumes from where it left off

## Features

- ✅ Automatic resume (skips already imported rows)
- ✅ Progress window with real-time status
- ✅ Automatic retry on network errors
- ✅ Status tracking in CSV file
- ✅ Excel-based configuration (no code changes needed)

## Need More Details?

See [TECHNICAL.md](TECHNICAL.md) for:
- Detailed field type detection
- Advanced configuration options
- Module function reference
- Detailed troubleshooting guide
- Environment variable setup
- Retry logic details
