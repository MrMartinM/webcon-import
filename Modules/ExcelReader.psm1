# ExcelReader.psm1
# Module for reading Excel files

function Read-WorkflowMapping {
    <#
    .SYNOPSIS
    Reads the Mapping-Workflow sheet from an Excel file and returns workflow configuration
    
    .PARAMETER FilePath
    Path to the Excel file
    
    .PARAMETER StartRow
    Row number to start reading from (default: 2, assumes row 1 is headers)
    
    .EXAMPLE
    $workflowConfig = Read-WorkflowMapping -FilePath "C:\data\workflows.xlsx"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [int]$StartRow = 2
    )
    
    # Check if ImportExcel module is available
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning "ImportExcel module not found. Installing..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
        Import-Module ImportExcel
    }
    
    if (-not (Test-Path $FilePath)) {
        throw "Excel file not found: $FilePath"
    }
    
    try {
        # Read the Mapping-Workflow sheet
        $data = Import-Excel -Path $FilePath -WorksheetName "Mapping-Workflow" -StartRow $StartRow
        
        if ($data.Count -eq 0) {
            throw "Mapping-Workflow sheet is empty or has no data rows"
        }
        
        # Get the first row (should contain the workflow configuration)
        $workflowRow = $data | Select-Object -First 1
        
        # Validate required columns
        $requiredColumns = @("WorkflowGuid", "FormTypeGuid")
        $missingColumns = $requiredColumns | Where-Object { -not $workflowRow.PSObject.Properties.Name -contains $_ }
        
        if ($missingColumns.Count -gt 0) {
            throw "Mapping-Workflow sheet is missing required columns: $($missingColumns -join ', ')"
        }
        
        # Build workflow config object
        $workflowConfig = @{
            WorkflowGuid = $workflowRow.WorkflowGuid
            FormTypeGuid = $workflowRow.FormTypeGuid
            Path = if ($workflowRow.PSObject.Properties.Name -contains "Path" -and $workflowRow.Path) { $workflowRow.Path } else { "default" }
            Mode = if ($workflowRow.PSObject.Properties.Name -contains "Mode" -and $workflowRow.Mode) { $workflowRow.Mode } else { "standard" }
        }
        
        Write-Verbose "Read workflow configuration from Mapping-Workflow sheet"
        return $workflowConfig
    }
    catch {
        Write-Error "Failed to read Mapping-Workflow sheet: $($_.Exception.Message)"
        throw
    }
}

function Read-MappingSheet {
    <#
    .SYNOPSIS
    Reads the Mapping-Fields sheet from an Excel file and returns field mappings
    
    .PARAMETER FilePath
    Path to the Excel file
    
    .PARAMETER StartRow
    Row number to start reading from (default: 2, assumes row 1 is headers)
    
    .EXAMPLE
    $mappings = Read-MappingSheet -FilePath "C:\data\workflows.xlsx"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [int]$StartRow = 2
    )
    
    # Check if ImportExcel module is available
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning "ImportExcel module not found. Installing..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
        Import-Module ImportExcel
    }
    
    if (-not (Test-Path $FilePath)) {
        throw "Excel file not found: $FilePath"
    }
    
    try {
        # Read the Mapping-Fields sheet
        $data = Import-Excel -Path $FilePath -WorksheetName "Mapping-Fields" -StartRow $StartRow
        
        # Validate required columns
        $requiredColumns = @("ExcelColumn", "FieldGuid", "FieldName", "FieldType")
        $firstRow = $data | Select-Object -First 1
        $missingColumns = $requiredColumns | Where-Object { -not $firstRow.PSObject.Properties.Name -contains $_ }
        
        if ($missingColumns.Count -gt 0) {
            throw "Mapping-Fields sheet is missing required columns: $($missingColumns -join ', ')"
        }
        
        Write-Verbose "Read $($data.Count) field mappings from Mapping-Fields sheet"
        return $data
    }
    catch {
        Write-Error "Failed to read Mapping-Fields sheet: $($_.Exception.Message)"
        throw
    }
}

function Read-ExcelFile {
    <#
    .SYNOPSIS
    Reads an Excel file and returns rows as objects
    
    .PARAMETER FilePath
    Path to the Excel file
    
    .PARAMETER WorksheetName
    Name of the worksheet to read (default: second sheet "Data")
    
    .PARAMETER StartRow
    Row number to start reading from (default: 2, assumes row 1 is headers)
    
    .EXAMPLE
    $rows = Read-ExcelFile -FilePath "C:\data\workflows.xlsx" -WorksheetName "Data"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [string]$WorksheetName = "Data",
        
        [Parameter(Mandatory=$false)]
        [int]$StartRow = 2
    )
    
    # Check if ImportExcel module is available
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning "ImportExcel module not found. Installing..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
        Import-Module ImportExcel
    }
    
    if (-not (Test-Path $FilePath)) {
        throw "Excel file not found: $FilePath"
    }
    
    try {
        $params = @{
            Path = $FilePath
            StartRow = $StartRow
            WorksheetName = $WorksheetName
        }
        
        $data = Import-Excel @params
        
        Write-Verbose "Read $($data.Count) rows from Excel file"
        return $data
    }
    catch {
        Write-Error "Failed to read Excel file: $($_.Exception.Message)"
        throw
    }
}

Export-ModuleMember -Function Read-WorkflowMapping, Read-MappingSheet, Read-ExcelFile

