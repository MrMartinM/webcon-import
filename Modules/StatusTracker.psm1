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
            $statusTable[$row.ID] = @{
                Status = $row.Status
                ImportedDate = $row.ImportedDate
                ErrorMessage = $row.ErrorMessage
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
        
        # Load existing status if file exists
        if (Test-Path $StatusFile) {
            $existingData = Import-Csv -Path $StatusFile
            foreach ($row in $existingData) {
                $statusTable[$row.ID] = @{
                    Status = $row.Status
                    ImportedDate = $row.ImportedDate
                    ErrorMessage = $row.ErrorMessage
                }
            }
        }
        
        # Update or add the status
        $importedDate = if ($Status -eq "Success") { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") } else { "" }
        
        $statusTable[$RowId] = @{
            Status = $Status
            ImportedDate = $importedDate
            ErrorMessage = $ErrorMessage
        }
        
        # Write back to CSV
        $statusDir = Split-Path -Path $StatusFile -Parent
        if ($statusDir -and -not (Test-Path $statusDir)) {
            New-Item -ItemType Directory -Path $statusDir -Force | Out-Null
        }
        
        $statusRecords = @()
        foreach ($id in $statusTable.Keys | Sort-Object) {
            $statusRecords += [PSCustomObject]@{
                ID = $id
                Status = $statusTable[$id].Status
                ImportedDate = $statusTable[$id].ImportedDate
                ErrorMessage = $statusTable[$id].ErrorMessage
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

Export-ModuleMember -Function Get-ImportStatus, Update-ImportStatus, IsRowImported

