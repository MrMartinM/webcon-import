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
  "Workflow": {
    "WorkflowGuid": "your-workflow-guid-here",
    "FormTypeGuid": "your-form-type-guid-here",
    "Path": "default",
    "Mode": "standard"
  },
  "Excel": {
    "FilePath": "C:\\data\\workflows.xlsx",
    "StartRow": 5
  },
  "ItemList": {
    "Enabled": false,
    "SheetName": "ItemList",
    "ItemListGuid": "your-item-list-guid-here",
    "ItemListName": "your-item-list-name-here"
  }
}
```

**Security**: Leave `ClientSecret` empty and set it as an environment variable:
```powershell
[System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-secret", "User")
```

**Important**: Make sure the system user (OAuth2 client) used by the script to call the API has **start workflow privileges** in Webcon.

### 4. Get Field Mappings from Webcon Database

Before preparing your Excel file, you need to get the field mapping information from your Webcon database using the provided stored procedures.

#### For Data Sheet (Workflow Fields)

1. **Create the stored procedure** in your Webcon database:
   ```sql
   -- Run the script: ImportWorkflows/SQL/mme_GetFieldDefinitions.sql
   ```

2. **Execute the stored procedure** with your workflow GUID:
   ```sql
   EXEC dbo.mme_GetFieldDefinitions @WF_Guid = 'your-workflow-guid-here'
   ```

3. **Copy the results** - The procedure returns:
   - Column headers: Friendly field names (from SQL column headers)
   - Row 1: DatabaseName (technical field names)
   - Row 2: Guid (field GUIDs)
   - Row 3: ColumnType (field type descriptions)

4. **Paste into Excel** - Copy these rows into your Excel "Data" sheet:
   - Row 1: Column headers (Friendly names from SQL column headers)
   - Row 2: DatabaseName values (from SQL row 1)
   - Row 3: Guid values (from SQL row 2)
   - Row 4: ColumnType values (from SQL row 3)

#### For ItemList Sheet (Item List Fields)

1. **Create the stored procedure** in your Webcon database:
   ```sql
   -- Run the script: ImportWorkflows/SQL/mme_GetFieldDetailDefinitions.sql
   ```

2. **Execute the stored procedure** with your item list configuration GUID:
   ```sql
   EXEC dbo.mme_GetFieldDetailDefinitions @WFCON_Guid = 'your-item-list-config-guid-here'
   ```

3. **Copy the results** - Same structure as above:
   - Row 1: Column headers (Friendly names from SQL column headers)
   - Row 2: DatabaseName values (from SQL row 1)
   - Row 3: Guid values (from SQL row 2)
   - Row 4: ColumnType values (from SQL row 3)

### 5. Prepare Excel File

Your Excel file needs **1 sheet** (plus optional ItemList sheet):

**Sheet: "Data"** - Field mappings and data in one sheet

**Rows 1-4** contain field metadata (populated from stored procedure results):
- **Row 1**: Friendly field names (e.g., "Active", "Company Name", "Email") - *From SQL column headers*
- **Row 2**: Database/Technical names (e.g., "WFD_AttBool1", "WFD_AttText1", "WFD_AttText2") - *From stored procedure row 1*
- **Row 3**: Field GUIDs - *From stored procedure row 2*
- **Row 4**: ColumnType (e.g., "Yes / No choice", "Single line of text", "Floating-point number") - *From stored procedure row 3*

**Row 5+**: Your actual data rows
| ID | Active | CompanyName | Email |
|----|--------|-------------|-------|
| 1  | Yes    | ACME Inc.   | info@acme.com |
| 2  | No     | Another Co. | contact@another.com |

**Optional: "ItemList" sheet** - If you're importing item lists, use the same structure (rows 1-4 = metadata, row 5+ = data). Get the mapping using `mme_GetFieldDetailDefinitions` stored procedure (see step 4 above). Make sure to include an `ID` column that matches the `ID` from the Data sheet to link item list rows to workflow instances.

### 6. Run
```powershell
cd ImportWorkflows
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
- Verify workflow and form type GUIDs in Config.json (Workflow section)
- Check that all required fields are mapped in rows 1-4 of the Data sheet

**Import stopped mid-process?**
- Just rerun the script - it automatically resumes from where it left off

## Features

- ✅ Automatic resume (skips already imported rows)
- ✅ Progress window with real-time status
- ✅ Automatic retry on network errors
- ✅ Status tracking in CSV file
- ✅ Excel-based configuration (no code changes needed)

## Need More Details?

See [ImportWorkflows/TECHNICAL.md](TECHNICAL.md) for:
- Detailed field type detection
- Advanced configuration options
- Module function reference
- Detailed troubleshooting guide
- Environment variable setup
- Retry logic details
