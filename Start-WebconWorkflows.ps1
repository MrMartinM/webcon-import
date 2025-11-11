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

# Step 2: Read workflow configuration from Excel Mapping-Workflow sheet
Write-Host "`nReading workflow configuration from Excel Mapping-Workflow sheet..." -ForegroundColor Yellow
$workflowConfig = Read-WorkflowMapping -FilePath $config.Excel.FilePath -StartRow $config.Excel.StartRow
Write-Host "Workflow Guid: $($workflowConfig.WorkflowGuid)" -ForegroundColor Green
Write-Host "Form Type Guid: $($workflowConfig.FormTypeGuid)" -ForegroundColor Green

# Step 3: Read field mappings from Excel Mapping-Fields sheet
Write-Host "`nReading field mappings from Excel Mapping-Fields sheet..." -ForegroundColor Yellow
$mappings = Read-MappingSheet -FilePath $config.Excel.FilePath -StartRow $config.Excel.StartRow
Write-Host "Found $($mappings.Count) field mappings" -ForegroundColor Green

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
$rows = Read-ExcelFile -FilePath $config.Excel.FilePath -WorksheetName "Data" -StartRow $config.Excel.StartRow
Write-Host "Found $($rows.Count) rows to process" -ForegroundColor Green

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
    $rowId = if ($row.PSObject.Properties.Name -contains "ID" -and $null -ne $row.ID) {
        $row.ID.ToString()
    } else {
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
                # Detect field type based on field name pattern
                $isChoiceField = $false
                $isBooleanField = $false
                $isDateTimeField = $false
                $isIntegerField = $false
                $isDecimalField = $false
                
                # Check if field name indicates a choice field (e.g., WFD_AttChoose2)
                if ($fieldMapping.FieldName -match "Choose" -or $fieldMapping.FieldName -match "Choice") {
                    $isChoiceField = $true
                    Write-Host "  Detected choice field: $($fieldMapping.FieldName)" -ForegroundColor Gray
                }
                
                # Check if field name indicates a boolean field (e.g., WFD_AttBool1)
                if ($fieldMapping.FieldName -match "AttBool") {
                    $isBooleanField = $true
                    Write-Host "  Detected boolean field: $($fieldMapping.FieldName)" -ForegroundColor Gray
                }
                
                # Check if field name indicates a datetime field (e.g., WFD_AttDateTime2)
                if ($fieldMapping.FieldName -match "AttDateTime") {
                    $isDateTimeField = $true
                    Write-Host "  Detected datetime field: $($fieldMapping.FieldName)" -ForegroundColor Gray
                }
                
                # Check if field name indicates an integer field (e.g., WFD_AttInt1)
                if ($fieldMapping.FieldName -match "AttInt") {
                    $isIntegerField = $true
                    Write-Host "  Detected integer field: $($fieldMapping.FieldName)" -ForegroundColor Gray
                }
                
                # Check if field name indicates a decimal field (e.g., WFD_AttDecimal7)
                if ($fieldMapping.FieldName -match "AttDecimal") {
                    $isDecimalField = $true
                    Write-Host "  Detected decimal field: $($fieldMapping.FieldName)" -ForegroundColor Gray
                }
                
                # Check for optional IsChoice column in mapping
                if ($fieldMapping.PSObject.Properties.Name -contains "IsChoice") {
                    $isChoiceColumn = $fieldMapping.IsChoice
                    if ($isChoiceColumn -eq "Yes" -or $isChoiceColumn -eq "True" -or $isChoiceColumn -eq $true -or $isChoiceColumn -eq 1) {
                        $isChoiceField = $true
                        Write-Host "  Explicitly marked as choice field: $($fieldMapping.FieldName)" -ForegroundColor Gray
                    }
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
                        Write-Host "  Choice field value: id='$choiceId', name='$choiceName'" -ForegroundColor Gray
                    } else {
                        # Single value - use as id, leave name blank
                        $choiceId = $excelValueStr.Trim()
                        $choiceName = ""
                        Write-Host "  Choice field value: id='$choiceId', name='' (blank)" -ForegroundColor Gray
                    }
                    
                    $fieldValue = @(
                        @{
                            id = $choiceId
                            name = $choiceName
                        }
                    )
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
                    Write-Host "  Regular field '$($fieldMapping.FieldName)': '$excelValueStr'" -ForegroundColor Gray
                }
                
                # Build form field with ordered properties
                $formField = [ordered]@{
                    guid     = $fieldMapping.FieldGuid
                    type     = $fieldMapping.FieldType
                    svalue   = $svalueStr
                    name     = $fieldMapping.FieldName
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
        
        # Start workflow with retry logic (one API call per row)
        $result = Start-WebconWorkflowWithRetry -BaseUrl $config.Webcon.BaseUrl `
                                                  -AccessToken $accessToken `
                                                  -DatabaseId $config.Webcon.DatabaseId `
                                                  -WorkflowGuid $workflowConfig.WorkflowGuid `
                                                  -FormTypeGuid $workflowConfig.FormTypeGuid `
                                                  -FormFields $formFields `
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

