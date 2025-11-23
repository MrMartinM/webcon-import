# StatusTracker.psm1
# Module for tracking import status via CSV file

function Get-ImportStatus {
    <#
    .SYNOPSIS
    Reads the status CSV file and returns a hashtable of row IDs and their status
    
    .PARAMETER StatusFile
    Path to the status CSV file
    
    .EXAMPLE
    $status = Get-ImportStatus -StatusFile "C:\data\workflows.status.csv"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$StatusFile
    )
    
    $statusTable = @{}
    
    # If file doesn't exist, return empty hashtable
    if (-not (Test-Path $StatusFile)) {
        Write-Verbose "Status file does not exist, returning empty status table"
        return $statusTable
    }
    
    try {
        $statusData = Import-Csv -Path $StatusFile
        
        foreach ($row in $statusData) {
            # Skip metadata rows (__START__ and __END__) when building status table
            if ($row.ID -ne "__START__" -and $row.ID -ne "__END__") {
                $statusTable[$row.ID] = @{
                    Status = $row.Status
                    ImportedDate = $row.ImportedDate
                    ErrorMessage = $row.ErrorMessage
                }
            }
        }
        
        Write-Verbose "Loaded $($statusTable.Count) status records from $StatusFile"
        return $statusTable
    }
    catch {
        Write-Warning "Failed to read status file: $($_.Exception.Message). Starting with empty status table."
        return $statusTable
    }
}

function Update-ImportStatus {
    <#
    .SYNOPSIS
    Updates or adds a status record for a row ID in the status CSV file
    
    .PARAMETER StatusFile
    Path to the status CSV file
    
    .PARAMETER RowId
    The row ID to update
    
    .PARAMETER Status
    Status value: "NotStarted", "Success", "Error"
    
    .PARAMETER ErrorMessage
    Optional error message if status is "Error"
    
    .EXAMPLE
    Update-ImportStatus -StatusFile "C:\data\workflows.status.csv" -RowId "1" -Status "Success"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$StatusFile,
        
        [Parameter(Mandatory=$true)]
        [string]$RowId,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("NotStarted", "Success", "Error")]
        [string]$Status,
        
        [Parameter(Mandatory=$false)]
        [string]$ErrorMessage = ""
    )
    
    try {
        $statusTable = @{}
        $statusRows = @()  # Array to preserve order of existing rows
        $startMetadata = $null
        $endMetadata = $null
        $isNewRow = $false
        
        # Load existing status if file exists
        if (Test-Path $StatusFile) {
            $existingData = Import-Csv -Path $StatusFile
            foreach ($row in $existingData) {
                if ($row.ID -eq "__START__") {
                    # Preserve start metadata
                    $startMetadata = @{
                        ID = "__START__"
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                    }
                }
                elseif ($row.ID -eq "__END__") {
                    # Preserve end metadata
                    $endMetadata = @{
                        ID = "__END__"
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                    }
                }
                else {
                    # Regular data row - preserve order and track in hashtable
                    $statusTable[$row.ID] = @{
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                        Order = $statusRows.Count  # Track original order
                    }
                    $statusRows += $row.ID
                }
            }
        }
        
        # Check if this is a new row or update to existing row
        $isNewRow = -not $statusTable.ContainsKey($RowId)
        
        # Update or add the status
        $importedDate = if ($Status -eq "Success") { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") } else { "" }
        
        if ($isNewRow) {
            # New row - will be appended at the end
            $statusTable[$RowId] = @{
                Status = $Status
                ImportedDate = $importedDate
                ErrorMessage = $ErrorMessage
                Order = [int]::MaxValue  # New rows get max order to append at end
            }
            $statusRows += $RowId
        }
        else {
            # Update existing row - preserve order
            $statusTable[$RowId].Status = $Status
            $statusTable[$RowId].ImportedDate = $importedDate
            $statusTable[$RowId].ErrorMessage = $ErrorMessage
        }
        
        # Write back to CSV
        $statusDir = Split-Path -Path $StatusFile -Parent
        if ($statusDir -and -not (Test-Path $statusDir)) {
            New-Item -ItemType Directory -Path $statusDir -Force | Out-Null
        }
        
        $statusRecords = @()
        
        # Add start metadata row first (if it exists)
        if ($startMetadata) {
            $statusRecords += [PSCustomObject]@{
                ID = $startMetadata.ID
                Status = $startMetadata.Status
                ImportedDate = $startMetadata.ImportedDate
                ErrorMessage = $startMetadata.ErrorMessage
            }
        }
        
        # Add data rows in preserved order (existing rows first, then new rows appended)
        foreach ($id in $statusRows) {
            if ($statusTable.ContainsKey($id)) {
                $statusRecords += [PSCustomObject]@{
                    ID = $id
                    Status = $statusTable[$id].Status
                    ImportedDate = $statusTable[$id].ImportedDate
                    ErrorMessage = $statusTable[$id].ErrorMessage
                }
            }
        }
        
        # Add end metadata row last (if it exists)
        if ($endMetadata) {
            $statusRecords += [PSCustomObject]@{
                ID = $endMetadata.ID
                Status = $endMetadata.Status
                ImportedDate = $endMetadata.ImportedDate
                ErrorMessage = $endMetadata.ErrorMessage
            }
        }
        
        $statusRecords | Export-Csv -Path $StatusFile -NoTypeInformation -Encoding UTF8
        
        Write-Verbose "Updated status for row $RowId to $Status"
    }
    catch {
        Write-Error "Failed to update status file: $($_.Exception.Message)"
        throw
    }
}

function IsRowImported {
    <#
    .SYNOPSIS
    Checks if a row ID has already been successfully imported
    
    .PARAMETER StatusTable
    Hashtable of status records (from Get-ImportStatus)
    
    .PARAMETER RowId
    The row ID to check
    
    .EXAMPLE
    if (IsRowImported -StatusTable $status -RowId "1") { Write-Host "Already imported" }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$StatusTable,
        
        [Parameter(Mandatory=$true)]
        [string]$RowId
    )
    
    if ($StatusTable.ContainsKey($RowId)) {
        return $StatusTable[$RowId].Status -eq "Success"
    }
    
    return $false
}

function Write-StartMetadata {
    <#
    .SYNOPSIS
    Writes or updates the start date-time metadata row in the status CSV file
    
    .PARAMETER StatusFile
    Path to the status CSV file
    
    .EXAMPLE
    Write-StartMetadata -StatusFile "C:\data\workflows.status.csv"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$StatusFile
    )
    
    try {
        $statusTable = @{}
        $statusRows = @()  # Array to preserve order of existing rows
        $startMetadata = $null
        $endMetadata = $null
        
        # Load existing status if file exists
        if (Test-Path $StatusFile) {
            $existingData = Import-Csv -Path $StatusFile
            foreach ($row in $existingData) {
                if ($row.ID -eq "__START__") {
                    $startMetadata = @{
                        ID = "__START__"
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                    }
                }
                elseif ($row.ID -eq "__END__") {
                    $endMetadata = @{
                        ID = "__END__"
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                    }
                }
                else {
                    $statusTable[$row.ID] = @{
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                    }
                    $statusRows += $row.ID
                }
            }
        }
        
        # Create or update start metadata
        $startDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $startMetadata = @{
            ID = "__START__"
            Status = "Metadata"
            ImportedDate = $startDate
            ErrorMessage = ""
        }
        
        # Write back to CSV
        $statusDir = Split-Path -Path $StatusFile -Parent
        if ($statusDir -and -not (Test-Path $statusDir)) {
            New-Item -ItemType Directory -Path $statusDir -Force | Out-Null
        }
        
        $statusRecords = @()
        
        # Add start metadata row first
        $statusRecords += [PSCustomObject]@{
            ID = $startMetadata.ID
            Status = $startMetadata.Status
            ImportedDate = $startMetadata.ImportedDate
            ErrorMessage = $startMetadata.ErrorMessage
        }
        
        # Add data rows in preserved order
        foreach ($id in $statusRows) {
            if ($statusTable.ContainsKey($id)) {
                $statusRecords += [PSCustomObject]@{
                    ID = $id
                    Status = $statusTable[$id].Status
                    ImportedDate = $statusTable[$id].ImportedDate
                    ErrorMessage = $statusTable[$id].ErrorMessage
                }
            }
        }
        
        # Add end metadata row last (if it exists)
        if ($endMetadata) {
            $statusRecords += [PSCustomObject]@{
                ID = $endMetadata.ID
                Status = $endMetadata.Status
                ImportedDate = $endMetadata.ImportedDate
                ErrorMessage = $endMetadata.ErrorMessage
            }
        }
        
        $statusRecords | Export-Csv -Path $StatusFile -NoTypeInformation -Encoding UTF8
        
        Write-Verbose "Written start metadata to $StatusFile"
    }
    catch {
        Write-Error "Failed to write start metadata: $($_.Exception.Message)"
        throw
    }
}

function Write-EndMetadata {
    <#
    .SYNOPSIS
    Writes or updates the end date-time metadata row in the status CSV file
    
    .PARAMETER StatusFile
    Path to the status CSV file
    
    .EXAMPLE
    Write-EndMetadata -StatusFile "C:\data\workflows.status.csv"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$StatusFile
    )
    
    try {
        $statusTable = @{}
        $statusRows = @()  # Array to preserve order of existing rows
        $startMetadata = $null
        $endMetadata = $null
        
        # Load existing status if file exists
        if (Test-Path $StatusFile) {
            $existingData = Import-Csv -Path $StatusFile
            foreach ($row in $existingData) {
                if ($row.ID -eq "__START__") {
                    $startMetadata = @{
                        ID = "__START__"
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                    }
                }
                elseif ($row.ID -eq "__END__") {
                    $endMetadata = @{
                        ID = "__END__"
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                    }
                }
                else {
                    $statusTable[$row.ID] = @{
                        Status = $row.Status
                        ImportedDate = $row.ImportedDate
                        ErrorMessage = $row.ErrorMessage
                    }
                    $statusRows += $row.ID
                }
            }
        }
        
        # Create or update end metadata
        $endDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $endMetadata = @{
            ID = "__END__"
            Status = "Metadata"
            ImportedDate = $endDate
            ErrorMessage = ""
        }
        
        # Write back to CSV
        $statusDir = Split-Path -Path $StatusFile -Parent
        if ($statusDir -and -not (Test-Path $statusDir)) {
            New-Item -ItemType Directory -Path $statusDir -Force | Out-Null
        }
        
        $statusRecords = @()
        
        # Add start metadata row first (if it exists)
        if ($startMetadata) {
            $statusRecords += [PSCustomObject]@{
                ID = $startMetadata.ID
                Status = $startMetadata.Status
                ImportedDate = $startMetadata.ImportedDate
                ErrorMessage = $startMetadata.ErrorMessage
            }
        }
        
        # Add data rows in preserved order
        foreach ($id in $statusRows) {
            if ($statusTable.ContainsKey($id)) {
                $statusRecords += [PSCustomObject]@{
                    ID = $id
                    Status = $statusTable[$id].Status
                    ImportedDate = $statusTable[$id].ImportedDate
                    ErrorMessage = $statusTable[$id].ErrorMessage
                }
            }
        }
        
        # Add end metadata row last
        $statusRecords += [PSCustomObject]@{
            ID = $endMetadata.ID
            Status = $endMetadata.Status
            ImportedDate = $endMetadata.ImportedDate
            ErrorMessage = $endMetadata.ErrorMessage
        }
        
        $statusRecords | Export-Csv -Path $StatusFile -NoTypeInformation -Encoding UTF8
        
        Write-Verbose "Written end metadata to $StatusFile"
    }
    catch {
        Write-Error "Failed to write end metadata: $($_.Exception.Message)"
        throw
    }
}

Export-ModuleMember -Function Get-ImportStatus, Update-ImportStatus, IsRowImported, Write-StartMetadata, Write-EndMetadata

