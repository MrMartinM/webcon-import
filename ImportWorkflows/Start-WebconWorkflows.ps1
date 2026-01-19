# Start-WebconWorkflows.ps1
# Main script to read Excel file and start Webcon workflows for each row

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\Config.json"
)

# Import modules
$modulePath = Join-Path $PSScriptRoot "Modules"
Import-Module (Join-Path $modulePath "ExcelReader.psm1") -Force
Import-Module (Join-Path $modulePath "WebconAPI.psm1") -Force
Import-Module (Join-Path $modulePath "StatusTracker.psm1") -Force
Import-Module (Join-Path $modulePath "ProgressWindow.psm1") -Force

# Load configuration
if (-not (Test-Path $ConfigPath)) {
    throw "Configuration file not found: $ConfigPath"
}

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

Write-Host "Starting Webcon workflow automation..." -ForegroundColor Green
Write-Host "Excel file: $($config.Excel.FilePath)" -ForegroundColor Cyan
Write-Host "Webcon URL: $($config.Webcon.BaseUrl)" -ForegroundColor Cyan

# Determine status file path
$statusFile = if ($config.StatusFile) {
    $config.StatusFile
} else {
    $excelDir = Split-Path -Path $config.Excel.FilePath -Parent
    $excelName = [System.IO.Path]::GetFileNameWithoutExtension($config.Excel.FilePath)
    Join-Path $excelDir "$excelName.status.csv"
}

Write-Host "Status file: $statusFile" -ForegroundColor Cyan

# Get retry settings
$maxRetries = if ($config.Retry.MaxRetries) { $config.Retry.MaxRetries } else { 3 }
$retryDelayBase = if ($config.Retry.RetryDelayBase) { $config.Retry.RetryDelayBase } else { 2 }

# Step 1: Load import status
Write-Host "`nLoading import status..." -ForegroundColor Yellow
$importStatus = Get-ImportStatus -StatusFile $statusFile
$alreadyImportedCount = ($importStatus.Values | Where-Object { $_.Status -eq "Success" }).Count
if ($alreadyImportedCount -gt 0) {
    Write-Host "Found $alreadyImportedCount already imported rows" -ForegroundColor Green
}

# Write start metadata to status file
Write-StartMetadata -StatusFile $statusFile

# Step 2: Read workflow configuration from Config.json
Write-Host "`nReading workflow configuration from Config.json..." -ForegroundColor Yellow
if (-not $config.Workflow) {
    throw "Workflow configuration not found in Config.json. Please add a 'Workflow' section with WorkflowGuid and FormTypeGuid."
}

$workflowConfig = @{
    WorkflowGuid = $config.Workflow.WorkflowGuid
    FormTypeGuid = $config.Workflow.FormTypeGuid
    Path = if ($config.Workflow.Path) { $config.Workflow.Path } else { "default" }
    Mode = if ($config.Workflow.Mode) { $config.Workflow.Mode } else { "standard" }
    BusinessEntityGuid = if ($config.Workflow.BusinessEntityGuid) { $config.Workflow.BusinessEntityGuid } else { "" }
}

Write-Host "Workflow Guid: $($workflowConfig.WorkflowGuid)" -ForegroundColor Green
Write-Host "Form Type Guid: $($workflowConfig.FormTypeGuid)" -ForegroundColor Green
if ($workflowConfig.BusinessEntityGuid -and $workflowConfig.BusinessEntityGuid.Trim() -ne "") {
    Write-Host "Business Entity Guid: $($workflowConfig.BusinessEntityGuid)" -ForegroundColor Green
}

# Step 3: Read field mappings from Data sheet (rows 1-4)
Write-Host "`nReading field mappings from Data sheet (rows 1-4)..." -ForegroundColor Yellow
$mappingResult = Read-FieldMappingsFromDataSheet -FilePath $config.Excel.FilePath -WorksheetName "Data"
$mappings = $mappingResult.Mappings
$dataIdColumnName = $mappingResult.IdColumnName
Write-Host "Found $($mappings.Count) field mappings" -ForegroundColor Green
Write-Host "Detected ID column name: '$dataIdColumnName'" -ForegroundColor Gray

# Step 3.5: Read item list configuration and data (if enabled)
$itemListEnabled = $false
$itemListMappings = @()
$itemListData = @()
$itemListGroupedById = @{}
$itemListGuid = $null
$itemListName = $null

# Check ItemList configuration
if ($config.ItemList) {
    Write-Host "`nItemList configuration found:" -ForegroundColor Cyan
    Write-Host "  Enabled: $($config.ItemList.Enabled)" -ForegroundColor Gray
    Write-Host "  SheetName: $(if ($config.ItemList.SheetName) { $config.ItemList.SheetName } else { 'ItemList (default)' })" -ForegroundColor Gray
    
    if ($config.ItemList.Enabled -eq $true) {
        $itemListEnabled = $true
        Write-Host "Item list import is enabled" -ForegroundColor Cyan
        
        # Get item list GUID and name from Config.json
        if (-not $config.ItemList.ItemListGuid -or -not $config.ItemList.ItemListName) {
            throw "ItemList.ItemListGuid and ItemList.ItemListName must be specified in Config.json when ItemList.Enabled is true"
        }
        $itemListGuid = $config.ItemList.ItemListGuid
        $itemListName = $config.ItemList.ItemListName
        Write-Host "  ItemListGuid: $itemListGuid" -ForegroundColor Gray
        Write-Host "  ItemListName: $itemListName" -ForegroundColor Gray
        
        # Read item list mappings from ItemList sheet (rows 1-4)
        $itemListSheetName = if ($config.ItemList.SheetName) { $config.ItemList.SheetName } else { "ItemList" }
        Write-Host "Reading item list mappings from Excel $itemListSheetName sheet (rows 1-4)..." -ForegroundColor Yellow
        try {
            $itemListMappingResult = Read-ItemListMappingsFromDataSheet -FilePath $config.Excel.FilePath -WorksheetName $itemListSheetName
            $itemListMappings = $itemListMappingResult.Mappings
            $itemListIdColumnName = $itemListMappingResult.IdColumnName
            Write-Host "Found $($itemListMappings.Count) item list column mappings" -ForegroundColor Green
            Write-Host "Detected ItemList ID column name: '$itemListIdColumnName'" -ForegroundColor Gray
            
            # Read item list data (starting from row 5)
            $startRow = if ($config.Excel.StartRow -and $config.Excel.StartRow -ge 5) { $config.Excel.StartRow } else { 5 }
            Write-Host "Reading item list data from Excel $itemListSheetName sheet (starting at row $startRow)..." -ForegroundColor Yellow
            $itemListData = Read-ItemListData -FilePath $config.Excel.FilePath -WorksheetName $itemListSheetName -StartRow $startRow -IdColumnName $itemListIdColumnName
            Write-Host "Found $($itemListData.Count) item list rows" -ForegroundColor Green
            
            # Group item list rows by ID using detected column name
            foreach ($itemListRow in $itemListData) {
                # Use bracket notation to access column with space or empty name
                $idValue = if ($itemListRow.PSObject.Properties.Name -contains $itemListIdColumnName) {
                    $itemListRow.$itemListIdColumnName
                } else {
                    # Try bracket notation as fallback
                    $itemListRow[$itemListIdColumnName]
                }
                
                $itemListRowId = if ($null -ne $idValue -and $idValue.ToString().Trim() -ne "") {
                    $idValue.ToString().Trim()
                } else {
                    ""
                }
                
                if ($itemListRowId -ne "") {
                    if (-not $itemListGroupedById.ContainsKey($itemListRowId)) {
                        $itemListGroupedById[$itemListRowId] = @()
                    }
                    $itemListGroupedById[$itemListRowId] += $itemListRow
                } else {
                    Write-Warning "Item list row has empty or missing ID (column '$itemListIdColumnName'), skipping: $($itemListRow | ConvertTo-Json -Compress)"
                }
            }
            Write-Host "Grouped item list rows into $($itemListGroupedById.Keys.Count) ID groups" -ForegroundColor Green
            
            # Debug: Show which IDs have item lists
            if ($itemListGroupedById.Keys.Count -gt 0) {
                Write-Host "Item list IDs found: $($itemListGroupedById.Keys -join ', ')" -ForegroundColor Gray
            } else {
                Write-Warning "No item list rows were grouped. Check that item list rows have valid ID values matching Data sheet IDs."
            }
        }
        catch {
            Write-Error "Failed to read item list configuration: $($_.Exception.Message)"
            Write-Error "Exception details: $($_.Exception | Format-List -Force | Out-String)"
            Write-Warning "Continuing without item lists."
            $itemListEnabled = $false
        }
    } else {
        Write-Host "Item list import is disabled" -ForegroundColor Gray
    }
} else {
    Write-Host "`nNo ItemList configuration found in Config.json" -ForegroundColor Gray
}

# Step 4: Get client secret from environment variable or config
$clientSecret = if ($config.Webcon.ClientSecret -and $config.Webcon.ClientSecret.Trim() -ne "") {
    $config.Webcon.ClientSecret
} else {
    # First try process environment variable (current session)
    $env:WEBCON_CLIENT_SECRET
}

# If not found in process environment, try reading from User registry (persistent environment variables)
if (-not $clientSecret -or $clientSecret.Trim() -eq "") {
    try {
        $userEnvKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("Environment")
        if ($userEnvKey) {
            $clientSecret = $userEnvKey.GetValue("WEBCON_CLIENT_SECRET", $null)
            $userEnvKey.Close()
        }
    }
    catch {
        # Ignore registry read errors
        Write-Verbose "Could not read WEBCON_CLIENT_SECRET from User registry: $($_.Exception.Message)"
    }
}

if (-not $clientSecret -or $clientSecret.Trim() -eq "") {
    throw "ClientSecret not found. Please set WEBCON_CLIENT_SECRET environment variable or provide it in Config.json"
}

# Step 5: Authenticate
Write-Host "`nAuthenticating with Webcon..." -ForegroundColor Yellow
$accessToken = Get-WebconToken -BaseUrl $config.Webcon.BaseUrl `
                                -ClientId $config.Webcon.ClientId `
                                -ClientSecret $clientSecret
Write-Host "Authentication successful!" -ForegroundColor Green

# Step 6: Read data from Excel Data sheet
Write-Host "`nReading data from Excel Data sheet..." -ForegroundColor Yellow
$startRow = if ($config.Excel.StartRow -and $config.Excel.StartRow -ge 5) { $config.Excel.StartRow } else { 5 }
Write-Host "Using StartRow: $startRow (will skip rows 2-4 as metadata)" -ForegroundColor Gray
$rows = Read-ExcelFile -FilePath $config.Excel.FilePath -WorksheetName "Data" -StartRow $startRow -IdColumnName $dataIdColumnName
Write-Host "Found $($rows.Count) rows to process" -ForegroundColor Green

# Debug: Show first few rows to verify they're not metadata
if ($rows.Count -gt 0) {
    Write-Host "`nFirst row preview (should be data, not metadata):" -ForegroundColor Gray
    $firstRow = $rows | Select-Object -First 1
    $firstRowProps = $firstRow.PSObject.Properties.Name | Select-Object -First 5
    foreach ($prop in $firstRowProps) {
        $propValue = if ($firstRow.PSObject.Properties.Name -contains $prop) { $firstRow.$prop } else { $firstRow[$prop] }
        Write-Host "  $prop = $propValue" -ForegroundColor Gray
    }
}

# Step 7: Process each row
$successCount = 0
$errorCount = 0
$skippedCount = 0
$errors = @()

# Calculate total rows to process (for progress tracking)
$totalRows = $rows.Count
$processedRows = 0

# Create and show progress window (only if there are rows to process)
$progressWindow = $null
if ($totalRows -gt 0) {
    $progressWindow = Show-ProgressWindow -TotalRows $totalRows
}

$rowIndex = 0
foreach ($row in $rows) {
    $rowIndex++
    
    # Generate row ID - check for ID column first, otherwise use row index
    # Use detected ID column name (may be space or empty string)
    $idValue = if ($row.PSObject.Properties.Name -contains $dataIdColumnName) {
        $row.$dataIdColumnName
    } else {
        # Try bracket notation as fallback
        $row[$dataIdColumnName]
    }
    
    $rowId = if ($null -ne $idValue -and $idValue.ToString().Trim() -ne "") {
        $idValue.ToString().Trim()
    } else {
        # Fallback to row number if ID column doesn't exist or is empty
        $rowIndex.ToString()
    }
    
    # Check for cancellation
    if ($progressWindow -and $progressWindow.IsCancelled()) {
        Write-Host "`nImport cancelled by user. Stopping at row $rowId..." -ForegroundColor Yellow
        break
    }
    
    # Check if row is already imported
    if (IsRowImported -StatusTable $importStatus -RowId $rowId) {
        $skippedCount++
        $processedRows++
        $progressPercent = [math]::Round(($processedRows / $totalRows) * 100, 1)
        
        # Update progress window
        if ($progressWindow) {
            $progressWindow.UpdateProgress($processedRows, "Row $rowId (Skipped)", $successCount, $errorCount, $skippedCount)
        }
        
        Write-Host "`nRow $rowId already imported, skipping... ($processedRows/$totalRows - $progressPercent%)" -ForegroundColor Gray
        continue
    }
    
    # Check for cancellation again before processing
    if ($progressWindow -and $progressWindow.IsCancelled()) {
        Write-Host "`nImport cancelled by user. Stopping at row $rowId..." -ForegroundColor Yellow
        break
    }
    
    Write-Host "`nProcessing row $rowId..." -ForegroundColor Yellow
    
    # Update progress window before processing
    if ($progressWindow) {
        $progressWindow.UpdateProgress($processedRows, "Row $rowId (Processing...)", $successCount, $errorCount, $skippedCount)
    }
    
    try {
        # Build form fields from Excel mappings
        $formFields = @()
        $fieldIndex = 0
        foreach ($fieldMapping in $mappings) {
            $fieldIndex++
            
            # Check for cancellation and process Windows messages periodically
            if ($progressWindow) {
                [System.Windows.Forms.Application]::DoEvents()
                if ($progressWindow.IsCancelled()) {
                    Write-Host "`nImport cancelled by user. Stopping during field mapping for row $rowId..." -ForegroundColor Yellow
                    # Break out of field mapping loop and skip to end of row processing
                    $formFields = @()  # Clear form fields since we're cancelling
                    break
                }
            }
            
            $excelValue = $row.($fieldMapping.ExcelColumn)
            
            if ($null -ne $excelValue) {
                # Detect field type - Primary: ColumnType parsing, Secondary: DatabaseName patterns
                $isChoiceField = $false
                $isBooleanField = $false
                $isDateTimeField = $false
                $isIntegerField = $false
                $isDecimalField = $false
                
                $columnType = if ($fieldMapping.ColumnType) { $fieldMapping.ColumnType.ToString().ToLower() } else { "" }
                $databaseName = if ($fieldMapping.DatabaseName) { $fieldMapping.DatabaseName.ToString() } else { "" }
                
                # Primary: Parse ColumnType strings
                if ($columnType -match "yes\s*/\s*no\s*choice") {
                    $isBooleanField = $true
                    Write-Host "  Detected boolean field (ColumnType: '$($fieldMapping.ColumnType)'): $databaseName" -ForegroundColor Gray
                }
                elseif ($columnType -match "floating-point\s*number") {
                    $isDecimalField = $true
                    Write-Host "  Detected decimal field (ColumnType: '$($fieldMapping.ColumnType)'): $databaseName" -ForegroundColor Gray
                }
                elseif ($columnType -match "choice" -and -not $columnType -match "yes\s*/\s*no") {
                    $isChoiceField = $true
                    Write-Host "  Detected choice field (ColumnType: '$($fieldMapping.ColumnType)'): $databaseName" -ForegroundColor Gray
                }
                # Secondary: Use DatabaseName patterns as fallback
                elseif ($databaseName -match "AttChoose" -or $databaseName -match "Choose") {
                    $isChoiceField = $true
                    Write-Host "  Detected choice field (DatabaseName pattern): $databaseName" -ForegroundColor Gray
                }
                elseif ($databaseName -match "AttBool") {
                    $isBooleanField = $true
                    Write-Host "  Detected boolean field (DatabaseName pattern): $databaseName" -ForegroundColor Gray
                }
                elseif ($databaseName -match "AttDateTime") {
                    $isDateTimeField = $true
                    Write-Host "  Detected datetime field (DatabaseName pattern): $databaseName" -ForegroundColor Gray
                }
                elseif ($databaseName -match "AttInt") {
                    $isIntegerField = $true
                    Write-Host "  Detected integer field (DatabaseName pattern): $databaseName" -ForegroundColor Gray
                }
                elseif ($databaseName -match "AttDecimal") {
                    $isDecimalField = $true
                    Write-Host "  Detected decimal field (DatabaseName pattern): $databaseName" -ForegroundColor Gray
                }
                
                # Build base form field structure with specific order: guid, type, svalue, name, formLayout, value
                # Convert to string and ensure proper UTF-8 encoding
                $excelValueStr = if ($null -ne $excelValue) { 
                    $excelValue.ToString() 
                } else { 
                    "" 
                }
                
                # Normalize string: clean and ensure valid encoding
                # Remove only problematic control characters that could break JSON serialization
                if ($excelValueStr.Length -gt 0) {
                    # Remove null bytes (these can cause JSON parsing errors)
                    $excelValueStr = $excelValueStr -replace "`0", ""
                    # Remove other control characters that could break JSON, but keep newlines, tabs, and carriage returns
                    # Keep: \n (0x0A), \r (0x0D), \t (0x09)
                    # Remove: other control chars (0x00-0x08, 0x0B-0x0C, 0x0E-0x1F, 0x7F)
                    $excelValueStr = $excelValueStr -replace "[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", ""
                }
                
                # Determine the value field based on field type
                $fieldValue = $null
                $svalueStr = $null  # Will be set based on field type
                
                if ($isChoiceField) {
                    # Handle choice fields - value must be an array
                    # Check if value is in format "id#name" or just a single value
                    if ($excelValueStr -match "^\s*([^#]+)\s*#\s*(.+)\s*$") {
                        # Format: "id#name"
                        $choiceId = $matches[1].Trim()
                        $choiceName = $matches[2].Trim()
                        
                        # Only include id if it's not empty
                        if ($choiceId -ne "") {
                            Write-Host "  Choice field value: id='$choiceId', name='$choiceName'" -ForegroundColor Gray
                            $choiceObj = @{
                                id = $choiceId
                                name = $choiceName
                            }
                        } else {
                            Write-Host "  Choice field value: name='$choiceName' (id is empty, not including)" -ForegroundColor Gray
                            $choiceObj = @{
                                name = $choiceName
                            }
                        }
                    } else {
                        # Single value - treat as name only (don't include id property)
                        $choiceName = $excelValueStr.Trim()
                        Write-Host "  Choice field value: name='$choiceName' (no id property)" -ForegroundColor Gray
                        
                        # Build choice object with only name (no id property)
                        $choiceObj = @{
                            name = $choiceName
                        }
                    }
                    
                    $fieldValue = @($choiceObj)
                    $svalueStr = $excelValueStr  # Keep original string for svalue
                }
                elseif ($isBooleanField) {
                    # Handle boolean fields - value must be boolean
                    if ($excelValue -is [bool]) {
                        $fieldValue = $excelValue
                    } elseif ($excelValueStr -match "^(true|1|yes|y)$" -or $excelValueStr -eq "1") {
                        $fieldValue = $true
                    } elseif ($excelValueStr -match "^(false|0|no|n)$" -or $excelValueStr -eq "0") {
                        $fieldValue = $false
                    } else {
                        # Try to parse as boolean
                        try {
                            $fieldValue = [System.Convert]::ToBoolean($excelValue)
                        } catch {
                            $fieldValue = $false
                        }
                    }
                    # For boolean fields, svalue should be empty string (API doesn't expect string representation)
                    $svalueStr = ""
                    Write-Host "  Boolean field value: $fieldValue" -ForegroundColor Gray
                }
                elseif ($isDateTimeField) {
                    # Handle datetime fields - value must be ISO 8601 string
                    try {
                        if ($excelValue -is [DateTime]) {
                            $fieldValue = $excelValue.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                        } else {
                            # Try to parse the string as DateTime
                            $dateTime = [DateTime]::Parse($excelValueStr)
                            $fieldValue = $dateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                        }
                        $svalueStr = $fieldValue  # Use ISO string for svalue
                        Write-Host "  DateTime field value: $fieldValue" -ForegroundColor Gray
                    } catch {
                        Write-Host "  Warning: Could not parse datetime value '$excelValueStr', using as string" -ForegroundColor Yellow
                        $fieldValue = $excelValueStr
                        $svalueStr = $excelValueStr
                    }
                }
                elseif ($isIntegerField) {
                    # Handle integer fields - value must be integer
                    try {
                        $fieldValue = [int]$excelValue
                        $svalueStr = $fieldValue.ToString()  # String representation for svalue
                        Write-Host "  Integer field value: $fieldValue" -ForegroundColor Gray
                    } catch {
                        Write-Host "  Warning: Could not parse integer value '$excelValueStr', using 0" -ForegroundColor Yellow
                        $fieldValue = 0
                        $svalueStr = "0"
                    }
                }
                elseif ($isDecimalField) {
                    # Handle decimal fields - value must be decimal
                    try {
                        $fieldValue = [decimal]$excelValue
                        $svalueStr = $fieldValue.ToString()  # String representation for svalue
                        Write-Host "  Decimal field value: $fieldValue" -ForegroundColor Gray
                    } catch {
                        Write-Host "  Warning: Could not parse decimal value '$excelValueStr', using 0" -ForegroundColor Yellow
                        $fieldValue = [decimal]0
                        $svalueStr = "0"
                    }
                }
                else {
                    # Regular field - value is a string
                    $fieldValue = $excelValueStr
                    $svalueStr = $excelValueStr
                    $fieldNameDisplay = if ($fieldMapping.DatabaseName) { $fieldMapping.DatabaseName } else { $fieldMapping.FieldName }
                    Write-Host "  Regular field '$fieldNameDisplay': '$excelValueStr'" -ForegroundColor Gray
                }
                
                # Build form field with ordered properties
                $formField = [ordered]@{
                    guid     = $fieldMapping.FieldGuid
                    type     = if ($fieldMapping.FieldType) { $fieldMapping.FieldType } else { "Unspecified" }
                    svalue   = $svalueStr
                    name     = if ($fieldMapping.DatabaseName) { $fieldMapping.DatabaseName } else { $fieldMapping.FieldName }
                    formLayout = @{
                        editability  = "Editable"
                        requiredness = "Optional"
                    }
                    value    = $fieldValue
                }
                
                $formFields += $formField
            }
        }
        
        # Check for cancellation before making API call
        if ($progressWindow -and $progressWindow.IsCancelled()) {
            Write-Host "`nImport cancelled by user. Stopping before processing row $rowId..." -ForegroundColor Yellow
            # Skip the API call and break out of row processing
            break
        }
        
        # Build item lists if enabled
        $itemLists = @()
        if ($itemListEnabled -and $itemListMappings.Count -gt 0) {
            # Use ID from current workflow row to find matching item list rows
            # Find matching item list rows
            if ($itemListGroupedById.ContainsKey($rowId)) {
                $matchingItemListRows = $itemListGroupedById[$rowId]
                Write-Host "  Found $($matchingItemListRows.Count) item list rows for ID $rowId" -ForegroundColor Gray
                
                # Build rows array
                $itemListRows = @()
                foreach ($itemListRow in $matchingItemListRows) {
                    $cells = @()
                    
                    foreach ($columnMapping in $itemListMappings) {
                        $excelValue = $itemListRow.($columnMapping.ExcelColumn)
                        
                        if ($null -ne $excelValue) {
                            # Detect field type - Primary: ColumnType parsing, Secondary: DatabaseName patterns (DET_ prefix)
                            $isChoiceField = $false
                            $isBooleanField = $false
                            $isDateTimeField = $false
                            $isIntegerField = $false
                            $isDecimalField = $false
                            $isLongTextField = $false
                            
                            $columnType = if ($columnMapping.ColumnType) { $columnMapping.ColumnType.ToString().ToLower() } else { "" }
                            $columnName = if ($columnMapping.DatabaseName) { $columnMapping.DatabaseName.ToString() } else { $columnMapping.ColumnName }
                            
                            # Primary: Parse ColumnType strings
                            if ($columnType -match "yes\s*/\s*no\s*choice") {
                                $isBooleanField = $true
                                Write-Host "    Detected boolean field (ColumnType: '$($columnMapping.ColumnType)'): $columnName" -ForegroundColor Gray
                            }
                            elseif ($columnType -match "floating-point\s*number") {
                                $isDecimalField = $true
                                Write-Host "    Detected decimal field (ColumnType: '$($columnMapping.ColumnType)'): $columnName" -ForegroundColor Gray
                            }
                            elseif ($columnType -match "multiple\s*lines\s*of\s*text") {
                                $isLongTextField = $true
                                Write-Host "    Detected long text field (ColumnType: '$($columnMapping.ColumnType)'): $columnName" -ForegroundColor Gray
                            }
                            elseif ($columnType -match "choice" -and -not $columnType -match "yes\s*/\s*no") {
                                $isChoiceField = $true
                                Write-Host "    Detected choice field (ColumnType: '$($columnMapping.ColumnType)'): $columnName" -ForegroundColor Gray
                            }
                            # Secondary: Use DatabaseName patterns as fallback
                            elseif ($columnName -match "^DET_Value") {
                                $isDecimalField = $true
                                Write-Host "    Detected decimal field (DatabaseName pattern): $columnName" -ForegroundColor Gray
                            }
                            elseif ($columnName -match "^DET_LongText") {
                                $isLongTextField = $true
                                Write-Host "    Detected long text field (DatabaseName pattern): $columnName" -ForegroundColor Gray
                            }
                            elseif ($columnName -match "^DET_Att") {
                                # Check if field name indicates a choice field
                                if ($columnName -match "Choose" -or $columnName -match "Choice") {
                                    $isChoiceField = $true
                                    Write-Host "    Detected choice field (DatabaseName pattern): $columnName" -ForegroundColor Gray
                                }
                                # Check if field name indicates a boolean field
                                elseif ($columnName -match "AttBool") {
                                    $isBooleanField = $true
                                    Write-Host "    Detected boolean field (DatabaseName pattern): $columnName" -ForegroundColor Gray
                                }
                                # Check if field name indicates a datetime field
                                elseif ($columnName -match "AttDateTime") {
                                    $isDateTimeField = $true
                                    Write-Host "    Detected datetime field (DatabaseName pattern): $columnName" -ForegroundColor Gray
                                }
                                # Check if field name indicates an integer field
                                elseif ($columnName -match "AttInt") {
                                    $isIntegerField = $true
                                    Write-Host "    Detected integer field (DatabaseName pattern): $columnName" -ForegroundColor Gray
                                }
                                # Default to string for other DET_Att* fields
                            }
                            
                            # Convert to string and ensure proper UTF-8 encoding
                            $excelValueStr = if ($null -ne $excelValue) { 
                                $excelValue.ToString() 
                            } else { 
                                "" 
                            }
                            
                            # Normalize string: clean and ensure valid encoding
                            if ($excelValueStr.Length -gt 0) {
                                $excelValueStr = $excelValueStr -replace "`0", ""
                                $excelValueStr = $excelValueStr -replace "[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", ""
                            }
                            
                            # Determine the value field based on field type
                            $cellValue = $null
                            $svalueStr = $null
                            
                            if ($isChoiceField) {
                                # Handle choice fields - value must be an array
                                if ($excelValueStr -match "^\s*([^#]+)\s*#\s*(.+)\s*$") {
                                    $choiceId = $matches[1].Trim()
                                    $choiceName = $matches[2].Trim()
                                    
                                    # Only include id if it's not empty
                                    if ($choiceId -ne "") {
                                        $choiceObj = @{
                                            id = $choiceId
                                            name = $choiceName
                                        }
                                    } else {
                                        $choiceObj = @{
                                            name = $choiceName
                                        }
                                    }
                                } else {
                                    # Single value - treat as name only (don't include id property)
                                    $choiceName = $excelValueStr.Trim()
                                    
                                    # Build choice object with only name (no id property)
                                    $choiceObj = @{
                                        name = $choiceName
                                    }
                                }
                                
                                $cellValue = @($choiceObj)
                                $svalueStr = $excelValueStr
                            }
                            elseif ($isBooleanField) {
                                # Handle boolean fields
                                if ($excelValue -is [bool]) {
                                    $cellValue = $excelValue
                                } elseif ($excelValueStr -match "^(true|1|yes|y)$" -or $excelValueStr -eq "1") {
                                    $cellValue = $true
                                } elseif ($excelValueStr -match "^(false|0|no|n)$" -or $excelValueStr -eq "0") {
                                    $cellValue = $false
                                } else {
                                    try {
                                        $cellValue = [System.Convert]::ToBoolean($excelValue)
                                    } catch {
                                        $cellValue = $false
                                    }
                                }
                                $svalueStr = ""
                            }
                            elseif ($isDateTimeField) {
                                # Handle datetime fields
                                try {
                                    if ($excelValue -is [DateTime]) {
                                        $cellValue = $excelValue.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                                    } else {
                                        $dateTime = [DateTime]::Parse($excelValueStr)
                                        $cellValue = $dateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                                    }
                                    $svalueStr = $cellValue
                                } catch {
                                    Write-Host "    Warning: Could not parse datetime value '$excelValueStr', using as string" -ForegroundColor Yellow
                                    $cellValue = $excelValueStr
                                    $svalueStr = $excelValueStr
                                }
                            }
                            elseif ($isIntegerField) {
                                # Handle integer fields
                                try {
                                    $cellValue = [int]$excelValue
                                    $svalueStr = $cellValue.ToString()
                                } catch {
                                    Write-Host "    Warning: Could not parse integer value '$excelValueStr', using 0" -ForegroundColor Yellow
                                    $cellValue = 0
                                    $svalueStr = "0"
                                }
                            }
                            elseif ($isDecimalField) {
                                # Handle decimal fields (DET_Value*)
                                try {
                                    $cellValue = [decimal]$excelValue
                                    $svalueStr = $cellValue.ToString()
                                } catch {
                                    Write-Host "    Warning: Could not parse decimal value '$excelValueStr', using 0" -ForegroundColor Yellow
                                    $cellValue = [decimal]0
                                    $svalueStr = "0"
                                }
                            }
                            else {
                                # Regular field or long text field - value is a string
                                $cellValue = $excelValueStr
                                $svalueStr = $excelValueStr
                            }
                            
                            # Build cell object (bare minimum: guid, svalue, value)
                            $cell = [ordered]@{
                                guid = $columnMapping.ColumnGuid
                                svalue = $svalueStr
                                value = $cellValue
                            }
                            
                            $cells += $cell
                        }
                    }
                    
                    # Build row object (only cells array)
                    $itemListRowObj = [ordered]@{
                        cells = $cells
                    }
                    
                    $itemListRows += $itemListRowObj
                }
                
                # Build item list object (bare minimum: guid, name, mode, rows)
                $itemList = [ordered]@{
                    guid = $itemListGuid
                    name = $itemListName
                    mode = "Incremental"
                    rows = $itemListRows
                }
                
                $itemLists = @($itemList)
                Write-Host "  Built item list with $($itemListRows.Count) rows" -ForegroundColor Gray
            } else {
                Write-Host "  No item list rows found for ID $rowId" -ForegroundColor Gray
            }
        }
        
        # Start workflow with retry logic (one API call per row)
        $result = Start-WebconWorkflowWithRetry -BaseUrl $config.Webcon.BaseUrl `
                                                  -AccessToken $accessToken `
                                                  -DatabaseId $config.Webcon.DatabaseId `
                                                  -WorkflowGuid $workflowConfig.WorkflowGuid `
                                                  -FormTypeGuid $workflowConfig.FormTypeGuid `
                                                  -BusinessEntityGuid $workflowConfig.BusinessEntityGuid `
                                                  -FormFields $formFields `
                                                  -ItemLists $itemLists `
                                                  -Path $workflowConfig.Path `
                                                  -Mode $workflowConfig.Mode `
                                                  -MaxRetries $maxRetries `
                                                  -RetryDelayBase $retryDelayBase
        
        # Update status to Success
        Update-ImportStatus -StatusFile $statusFile -RowId $rowId -Status "Success"
        
        $processedRows++
        $successCount++
        $progressPercent = [math]::Round(($processedRows / $totalRows) * 100, 1)
        
        # Update progress window
        if ($progressWindow) {
            $progressWindow.UpdateProgress($processedRows, "Row $rowId (Success)", $successCount, $errorCount, $skippedCount)
        }
        
        Write-Host "Workflow started successfully! ($processedRows/$totalRows - $progressPercent%)" -ForegroundColor Green
    }
    catch {
        $errorMsg = $_.Exception.Message
        
        $processedRows++
        $errorCount++
        $progressPercent = [math]::Round(($processedRows / $totalRows) * 100, 1)
        
        # Update progress window
        if ($progressWindow) {
            $progressWindow.UpdateProgress($processedRows, "Row $rowId (Error)", $successCount, $errorCount, $skippedCount)
        }
        
        Write-Host "Error processing row $rowId : $errorMsg ($processedRows/$totalRows - $progressPercent%)" -ForegroundColor Red
        
        # Update status to Error
        Update-ImportStatus -StatusFile $statusFile -RowId $rowId -Status "Error" -ErrorMessage $errorMsg
        
        $errors += @{
            RowId = $rowId
            Row = $row
            Error = $errorMsg
        }
    }
    
    # Check for cancellation after processing each row
    if ($progressWindow -and $progressWindow.IsCancelled()) {
        Write-Host "`nImport cancelled by user. Stopping after row $rowId..." -ForegroundColor Yellow
        break
    }
}

# Check if cancelled and handle accordingly
$wasCancelled = $false
if ($progressWindow) {
    $wasCancelled = $progressWindow.IsCancelled()
    if ($wasCancelled) {
        $progressWindow.SetCancelled()
    } else {
        $progressWindow.Close()
    }
}

# Summary
if ($wasCancelled) {
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "Import Cancelled!" -ForegroundColor Yellow
    Write-Host "Processed before cancellation: $processedRows / $totalRows" -ForegroundColor Yellow
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Skipped (already imported): $skippedCount" -ForegroundColor $(if ($skippedCount -gt 0) { "Yellow" } else { "Gray" })
    Write-Host "Errors: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "`nYou can rerun the script to continue from where it left off." -ForegroundColor Gray
} else {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Processing Complete!" -ForegroundColor Green
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Skipped (already imported): $skippedCount" -ForegroundColor $(if ($skippedCount -gt 0) { "Yellow" } else { "Gray" })
    Write-Host "Errors: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
    Write-Host "Total processed: $($successCount + $errorCount)" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}

if ($errors.Count -gt 0) {
    Write-Host "`nErrors:" -ForegroundColor Red
    $errors | ForEach-Object {
        Write-Host "  Row $($_.RowId): $($_.Error)" -ForegroundColor Red
    }
}

Write-Host "`nStatus file: $statusFile" -ForegroundColor Cyan
if (-not $wasCancelled) {
    Write-Host "You can rerun this script to retry failed rows or continue from where you left off." -ForegroundColor Gray
}

# Write end metadata to status file
Write-EndMetadata -StatusFile $statusFile

# Wait for user to close progress window
if ($progressWindow -and $progressWindow.Form) {
    try {
        # Check if form is still valid and not disposed
        if (-not $progressWindow.Form.IsDisposed) {
            Write-Host "`nProgress window is open. Close it when done reviewing." -ForegroundColor Cyan
            # Hide the form first since it's already visible (from Show()), then show as modal dialog
            # This prevents the "Form that is already visible cannot be displayed as a modal dialog" error
            if ($progressWindow.Form.Visible) {
                $progressWindow.Form.Hide()
            }
            # Show as modal dialog - this will block until user closes the window
            $progressWindow.Form.ShowDialog() | Out-Null
        }
    }
    catch {
        # If form was already closed, disposed, or other error occurred, just continue silently
        # This can happen if user manually closed the window before script reached this point
        Write-Verbose "Progress window handling: $($_.Exception.Message)"
    }
    finally {
        # Ensure form is properly disposed
        try {
            if ($progressWindow.Form -and -not $progressWindow.Form.IsDisposed) {
                $progressWindow.Form.Dispose()
            }
        }
        catch {
            # Ignore disposal errors
            Write-Verbose "Form disposal: $($_.Exception.Message)"
        }
    }
}

