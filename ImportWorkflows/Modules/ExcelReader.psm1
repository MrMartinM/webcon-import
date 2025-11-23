# ExcelReader.psm1
# Module for reading Excel files

function Read-FieldMappingsFromDataSheet {
    <#
    .SYNOPSIS
    Reads field mappings from rows 1-4 of the Data sheet in an Excel file
    
    .PARAMETER FilePath
    Path to the Excel file
    
    .PARAMETER WorksheetName
    Name of the worksheet to read (default: "Data")
    
    .EXAMPLE
    $mappings = Read-FieldMappingsFromDataSheet -FilePath "C:\data\workflows.xlsx"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [string]$WorksheetName = "Data"
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
        # Read rows 1-4 as metadata (column header + 3 data rows from SQL)
        # Row 1: Column headers (Friendly names from SQL column headers)
        # Row 2: Database/Technical names (from SQL row 1)
        # Row 3: GUIDs (from SQL row 2)
        # Row 4: ColumnType (from SQL row 3)
        
        # Read rows 1-4: Row 1 has headers, rows 2-4 have data
        # Use StartRow=1 to read with row 1 as headers, then access rows 2-4
        $allRows = Import-Excel -Path $FilePath -WorksheetName $WorksheetName -StartRow 1
        
        if ($null -eq $allRows -or $allRows.Count -eq 0) {
            throw "Data sheet is missing metadata rows (rows 1-4)"
        }
        
        # Get column names from first row (which becomes the property names)
        $firstDataRow = $allRows | Select-Object -First 1
        $columns = $firstDataRow.PSObject.Properties.Name | Where-Object { $_ -ne "RowType" }
        
        if ($columns.Count -eq 0) {
            throw "No field columns found in Data sheet"
        }
        
        # Row 1 (index 0) = headers (already used as property names)
        # Row 2 (index 1) = DatabaseName
        # Row 3 (index 2) = Guid
        # Row 4 (index 3) = ColumnType
        
        $row2Obj = if ($allRows.Count -gt 0) { $allRows[0] } else { $null }
        $row3Obj = if ($allRows.Count -gt 1) { $allRows[1] } else { $null }
        $row4Obj = if ($allRows.Count -gt 2) { $allRows[2] } else { $null }
        
        # Detect ID column by checking row 4 (ColumnType row) for "ID" value
        $idColumnName = $null
        foreach ($col in $columns) {
            $columnType = if ($row4Obj -and $row4Obj.$col) { 
                $row4Obj.$col.ToString().Trim() 
            } else { "" }
            
            if ($columnType -eq "ID") {
                $idColumnName = $col
                Write-Verbose "Detected ID column: '$idColumnName' (header is space or empty)"
                break
            }
        }
        
        if (-not $idColumnName) {
            Write-Warning "ID column not found in $WorksheetName sheet (no column with 'ID' in row 4). Will use first column as fallback."
            $idColumnName = $columns[0]
        }
        
        # Build mappings array
        $mappings = @()
        foreach ($col in $columns) {
            $friendlyName = $col  # Column header is the friendly name
            
            $databaseName = if ($row2Obj -and $row2Obj.$col) { 
                $row2Obj.$col.ToString() 
            } else { "" }
            
            $fieldGuid = if ($row3Obj -and $row3Obj.$col) { 
                $row3Obj.$col.ToString() 
            } else { "" }
            
            $columnType = if ($row4Obj -and $row4Obj.$col) { 
                $row4Obj.$col.ToString() 
            } else { "" }
            
            # Only add mapping if we have essential information
            if ($fieldGuid -and $databaseName) {
                $mapping = [PSCustomObject]@{
                    ExcelColumn = $col
                    FriendlyName = $friendlyName
                    DatabaseName = $databaseName
                    FieldGuid = $fieldGuid
                    FieldName = $databaseName  # Use DatabaseName as FieldName for compatibility
                    DataType = ""  # Not used, kept for compatibility
                    ColumnType = $columnType
                    FieldType = "Unspecified"  # Default, will be determined by type detection
                }
                $mappings += $mapping
            }
        }
        
        Write-Verbose "Read $($mappings.Count) field mappings from $WorksheetName sheet (rows 1-4)"
        
        # Return both mappings and ID column name
        return @{
            Mappings = $mappings
            IdColumnName = $idColumnName
        }
    }
    catch {
        Write-Error "Failed to read field mappings from $WorksheetName sheet: $($_.Exception.Message)"
        throw
    }
}

function Read-ExcelFile {
    <#
    .SYNOPSIS
    Reads an Excel file and returns rows as objects (data starts at row 5, rows 1-4 are metadata)
    
    .PARAMETER FilePath
    Path to the Excel file
    
    .PARAMETER WorksheetName
    Name of the worksheet to read (default: "Data")
    
    .PARAMETER StartRow
    Row number to start reading data from (default: 5, rows 1-4 are metadata)
    
    .PARAMETER IdColumnName
    Name of the ID column (can be space or empty string)
    
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
        [int]$StartRow = 5,
        
        [Parameter(Mandatory=$false)]
        [string]$IdColumnName = "ID"
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
        # Import-Excel StartRow parameter uses that row as headers
        # We need row 1 as headers, but skip rows 2-4 (metadata)
        # So we use StartRow=1 to get correct headers, then filter out metadata rows
        $params = @{
            Path = $FilePath
            StartRow = 1  # Always use row 1 as headers
            WorksheetName = $WorksheetName
        }
        
        $allData = Import-Excel @params
        
        # Filter out metadata rows (rows 2-4, which are indices 0-2 in the array)
        # When StartRow=1, Import-Excel reads row 1 as headers and rows 2+ as data
        # So: index 0 = Excel row 2 (DatabaseName), index 1 = Excel row 3 (Guid), index 2 = Excel row 4 (ColumnType)
        # We always skip exactly 3 rows (rows 2-4) since metadata is always in rows 2-4
        
        Write-Verbose "Total rows read from Excel: $($allData.Count)"
        
        # Function to check if a row looks like metadata
        function Test-IsMetadataRow {
            param($row)
            if (-not $row) { return $false }
            
            $guidCount = 0
            $fieldNameCount = 0
            $totalProps = 0
            
            foreach ($prop in $row.PSObject.Properties) {
                $value = if ($prop.Value) { $prop.Value.ToString().Trim() } else { "" }
                if ($value -eq "") { continue }
                
                $totalProps++
                
                # Check for GUID pattern (8-4-4-4-12 hex digits with hyphens)
                if ($value -match "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$") {
                    $guidCount++
                }
                # Check for field name patterns (WFD_ or DET_ prefix)
                if ($value -match "^(WFD_|DET_)") {
                    $fieldNameCount++
                }
            }
            
            # If more than 50% of non-empty values are GUIDs or field names, it's likely metadata
            if ($totalProps -gt 0) {
                $metadataRatio = ($guidCount + $fieldNameCount) / $totalProps
                return $metadataRatio -gt 0.5
            }
            return $false
        }
        
        # Skip first 3 rows (should be metadata), but also check dynamically
        $data = @()
        $skippedCount = 0
        
        for ($i = 0; $i -lt $allData.Count; $i++) {
            $row = $allData[$i]
            
            # Always skip first 3 rows (indices 0-2, which are Excel rows 2-4)
            if ($i -lt 3) {
                $skippedCount++
                Write-Verbose "Skipping row $($i + 2) (metadata row $($i + 1))"
                continue
            }
            
            # Also check if this row looks like metadata (safety check)
            if (Test-IsMetadataRow -row $row) {
                Write-Warning "Row $($i + 2) appears to be metadata (contains GUIDs or field names). Skipping."
                $skippedCount++
                continue
            }
            
            # This is a data row
            $data += $row
        }
        
        Write-Verbose "Skipped $skippedCount metadata rows, kept $($data.Count) data rows"
        
        if ($data.Count -eq 0 -and $allData.Count -gt 0) {
            Write-Warning "No data rows found after filtering metadata. Check Excel file structure."
        }
        
        Write-Verbose "Read $($data.Count) data rows from $WorksheetName sheet (starting at row 5, skipped $skippedCount metadata rows)"
        return $data
    }
    catch {
        Write-Error "Failed to read Excel file: $($_.Exception.Message)"
        throw
    }
}

function Read-ItemListMappingsFromDataSheet {
    <#
    .SYNOPSIS
    Reads item list column mappings from rows 1-4 of the ItemList sheet in an Excel file
    
    .PARAMETER FilePath
    Path to the Excel file
    
    .PARAMETER WorksheetName
    Name of the worksheet to read (default: "ItemList")
    
    .EXAMPLE
    $itemListMappings = Read-ItemListMappingsFromDataSheet -FilePath "C:\data\workflows.xlsx" -WorksheetName "ItemList"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [string]$WorksheetName = "ItemList"
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
        # Read rows 1-4 as metadata (column header + 3 data rows from SQL)
        # Row 1: Column headers (Friendly names from SQL column headers)
        # Row 2: Database/Technical names (from SQL row 1)
        # Row 3: GUIDs (from SQL row 2)
        # Row 4: DataType (manual entry, inferred from DatabaseName patterns if not provided)
        
        # Read rows 1-4: Row 1 has headers, rows 2-4 have data
        # Use StartRow=1 to read with row 1 as headers, then access rows 2-4
        $allRows = Import-Excel -Path $FilePath -WorksheetName $WorksheetName -StartRow 1
        
        if ($null -eq $allRows -or $allRows.Count -eq 0) {
            throw "ItemList sheet is missing metadata rows (rows 1-4)"
        }
        
        # Get column names from first row (which becomes the property names)
        $firstDataRow = $allRows | Select-Object -First 1
        $columns = $firstDataRow.PSObject.Properties.Name | Where-Object { $_ -ne "RowType" }
        
        if ($columns.Count -eq 0) {
            throw "No item list columns found in $WorksheetName sheet"
        }
        
        # Row 1 (index 0) = headers (already used as property names)
        # Row 2 (index 1) = DatabaseName
        # Row 3 (index 2) = Guid
        # Row 4 (index 3) = ColumnType
        
        $row2Obj = if ($allRows.Count -gt 0) { $allRows[0] } else { $null }
        $row3Obj = if ($allRows.Count -gt 1) { $allRows[1] } else { $null }
        $row4Obj = if ($allRows.Count -gt 2) { $allRows[2] } else { $null }
        
        # Detect ID column by checking row 4 (ColumnType row) for "ID" value
        $idColumnName = $null
        foreach ($col in $columns) {
            $columnType = if ($row4Obj -and $row4Obj.$col) { 
                $row4Obj.$col.ToString().Trim() 
            } else { "" }
            
            if ($columnType -eq "ID") {
                $idColumnName = $col
                Write-Verbose "Detected ID column: '$idColumnName' (header is space or empty)"
                break
            }
        }
        
        if (-not $idColumnName) {
            Write-Warning "ID column not found in $WorksheetName sheet (no column with 'ID' in row 4). Will use first column as fallback."
            $idColumnName = $columns[0]
        }
        
        # Build mappings array
        $mappings = @()
        foreach ($col in $columns) {
            $friendlyName = $col  # Column header is the friendly name
            
            $databaseName = if ($row2Obj -and $row2Obj.$col) { 
                $row2Obj.$col.ToString() 
            } else { "" }
            
            $columnGuid = if ($row3Obj -and $row3Obj.$col) { 
                $row3Obj.$col.ToString() 
            } else { "" }
            
            $columnType = if ($row4Obj -and $row4Obj.$col) { 
                $row4Obj.$col.ToString() 
            } else { "" }
            
            # Only add mapping if we have essential information
            if ($columnGuid -and $databaseName) {
                $mapping = [PSCustomObject]@{
                    ExcelColumn = $col
                    FriendlyName = $friendlyName
                    DatabaseName = $databaseName
                    ColumnGuid = $columnGuid
                    ColumnName = $databaseName  # Use DatabaseName as ColumnName for compatibility
                    DataType = ""  # Not used, kept for compatibility
                    ColumnType = $columnType
                }
                $mappings += $mapping
            }
        }
        
        Write-Verbose "Read $($mappings.Count) item list column mappings from $WorksheetName sheet (rows 1-4)"
        
        # Return both mappings and ID column name
        return @{
            Mappings = $mappings
            IdColumnName = $idColumnName
        }
    }
    catch {
        Write-Error "Failed to read item list mappings from $WorksheetName sheet: $($_.Exception.Message)"
        throw
    }
}

function Read-ItemListData {
    <#
    .SYNOPSIS
    Reads item list data from an Excel sheet and returns rows as objects (data starts at row 5, rows 1-4 are metadata)
    
    .PARAMETER FilePath
    Path to the Excel file
    
    .PARAMETER WorksheetName
    Name of the worksheet to read (default: "ItemList")
    
    .PARAMETER StartRow
    Row number to start reading data from (default: 5, rows 1-4 are metadata)
    
    .PARAMETER IdColumnName
    Name of the ID column (can be space or empty string)
    
    .EXAMPLE
    $itemListRows = Read-ItemListData -FilePath "C:\data\workflows.xlsx" -WorksheetName "ItemList"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [string]$WorksheetName = "ItemList",
        
        [Parameter(Mandatory=$false)]
        [int]$StartRow = 5,
        
        [Parameter(Mandatory=$false)]
        [string]$IdColumnName = "ID"
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
        # Import-Excel StartRow parameter uses that row as headers
        # We need row 1 as headers, but skip rows 2-4 (metadata)
        # So we use StartRow=1 to get correct headers, then filter out metadata rows
        $params = @{
            Path = $FilePath
            StartRow = 1  # Always use row 1 as headers
            WorksheetName = $WorksheetName
        }
        
        $allData = Import-Excel @params
        
        # Filter out metadata rows (rows 2-4, which are indices 0-2 in the array)
        # When StartRow=1, Import-Excel reads row 1 as headers and rows 2+ as data
        # So: index 0 = Excel row 2 (DatabaseName), index 1 = Excel row 3 (Guid), index 2 = Excel row 4 (ColumnType)
        # We always skip exactly 3 rows (rows 2-4) since metadata is always in rows 2-4
        
        Write-Verbose "Total rows read from Excel: $($allData.Count)"
        
        # Function to check if a row looks like metadata
        function Test-IsMetadataRow {
            param($row)
            if (-not $row) { return $false }
            
            $guidCount = 0
            $fieldNameCount = 0
            $totalProps = 0
            
            foreach ($prop in $row.PSObject.Properties) {
                $value = if ($prop.Value) { $prop.Value.ToString().Trim() } else { "" }
                if ($value -eq "") { continue }
                
                $totalProps++
                
                # Check for GUID pattern (8-4-4-4-12 hex digits with hyphens)
                if ($value -match "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$") {
                    $guidCount++
                }
                # Check for field name patterns (WFD_ or DET_ prefix)
                if ($value -match "^(WFD_|DET_)") {
                    $fieldNameCount++
                }
            }
            
            # If more than 50% of non-empty values are GUIDs or field names, it's likely metadata
            if ($totalProps -gt 0) {
                $metadataRatio = ($guidCount + $fieldNameCount) / $totalProps
                return $metadataRatio -gt 0.5
            }
            return $false
        }
        
        # Skip first 3 rows (should be metadata), but also check dynamically
        $data = @()
        $skippedCount = 0
        
        for ($i = 0; $i -lt $allData.Count; $i++) {
            $row = $allData[$i]
            
            # Always skip first 3 rows (indices 0-2, which are Excel rows 2-4)
            if ($i -lt 3) {
                $skippedCount++
                Write-Verbose "Skipping row $($i + 2) (metadata row $($i + 1))"
                continue
            }
            
            # Also check if this row looks like metadata (safety check)
            if (Test-IsMetadataRow -row $row) {
                Write-Warning "Row $($i + 2) appears to be metadata (contains GUIDs or field names). Skipping."
                $skippedCount++
                continue
            }
            
            # This is a data row
            $data += $row
        }
        
        Write-Verbose "Skipped $skippedCount metadata rows, kept $($data.Count) data rows"
        
        if ($data.Count -eq 0 -and $allData.Count -gt 0) {
            Write-Warning "No data rows found after filtering metadata. Check Excel file structure."
        }
        
        # Validate that ID column exists (use detected column name, which may be space or empty)
        if ($data.Count -gt 0) {
            $firstRow = $data | Select-Object -First 1
            $hasIdColumn = $false
            
            # Check if ID column exists (handle space or empty string column names)
            foreach ($propName in $firstRow.PSObject.Properties.Name) {
                if ($propName -eq $IdColumnName) {
                    $hasIdColumn = $true
                    break
                }
            }
            
            if (-not $hasIdColumn) {
                $availableColumns = $firstRow.PSObject.Properties.Name -join ", "
                throw "$WorksheetName sheet is missing required ID column '$IdColumnName'. Available columns: $availableColumns"
            }
        }
        
        Write-Verbose "Read $($data.Count) item list data rows from $WorksheetName sheet (starting at row 5, skipped $skippedCount metadata rows)"
        return $data
    }
    catch {
        Write-Error "Failed to read item list data from $WorksheetName sheet: $($_.Exception.Message)"
        throw
    }
}

Export-ModuleMember -Function Read-FieldMappingsFromDataSheet, Read-ExcelFile, Read-ItemListMappingsFromDataSheet, Read-ItemListData

