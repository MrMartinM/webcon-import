# Add-WebconAttachment.ps1
# Script to upload a file as an attachment to a Webcon element

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,
    
    [Parameter(Mandatory=$true)]
    [int]$ElementId,
    
    [Parameter(Mandatory=$false)]
    [string]$Name,
    
    [Parameter(Mandatory=$false)]
    [string]$Description = "",
    
    [Parameter(Mandatory=$false)]
    [string]$Group = "",
    
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\Config.json"
)

# Import modules
$modulePath = Join-Path $PSScriptRoot "Modules"
Import-Module (Join-Path $modulePath "WebconAPI.psm1") -Force

# Load configuration
if (-not (Test-Path $ConfigPath)) {
    throw "Configuration file not found: $ConfigPath"
}

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

Write-Host "Uploading attachment to Webcon element..." -ForegroundColor Green
Write-Host "File: $FilePath" -ForegroundColor Cyan
Write-Host "Element ID: $ElementId" -ForegroundColor Cyan
Write-Host "Webcon URL: $($config.Webcon.BaseUrl)" -ForegroundColor Cyan

# Validate file exists
if (-not (Test-Path $FilePath)) {
    throw "File not found: $FilePath"
}

# Get file info
$fileInfo = Get-Item $FilePath
$fileName = if ($Name) { $Name } else { $fileInfo.Name }
$fileSize = $fileInfo.Length

Write-Host "File name: $fileName" -ForegroundColor Gray
Write-Host "File size: $fileSize bytes" -ForegroundColor Gray

# Read file bytes and convert to base64
Write-Host "`nReading file and converting to base64..." -ForegroundColor Yellow
try {
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $base64Content = [System.Convert]::ToBase64String($fileBytes)
    $base64Length = $base64Content.Length
    Write-Host "Base64 length: $base64Length characters" -ForegroundColor Gray
}
catch {
    Write-Error "Failed to read file: $($_.Exception.Message)"
    throw
}

# Get client secret from environment variable or config
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

# Authenticate
Write-Host "`nAuthenticating with Webcon..." -ForegroundColor Yellow
try {
    $accessToken = Get-WebconToken -BaseUrl $config.Webcon.BaseUrl `
                                    -ClientId $config.Webcon.ClientId `
                                    -ClientSecret $clientSecret
    Write-Host "Authentication successful!" -ForegroundColor Green
}
catch {
    Write-Error "Failed to authenticate: $($_.Exception.Message)"
    throw
}

# Upload attachment with retry logic
Write-Host "`nUploading attachment..." -ForegroundColor Yellow
try {
    # Get retry settings from config (if available)
    $maxRetries = if ($config.Retry -and $config.Retry.MaxRetries) { $config.Retry.MaxRetries } else { 3 }
    $retryDelayBase = if ($config.Retry -and $config.Retry.RetryDelayBase) { $config.Retry.RetryDelayBase } else { 2 }
    
    $result = Add-WebconAttachmentWithRetry -BaseUrl $config.Webcon.BaseUrl `
                                           -AccessToken $accessToken `
                                           -DatabaseId $config.Webcon.DatabaseId `
                                           -ElementId $ElementId `
                                           -Name $fileName `
                                           -Description $Description `
                                           -Group $Group `
                                           -Content $base64Content `
                                           -MaxRetries $maxRetries `
                                           -RetryDelayBase $retryDelayBase
    
    Write-Host "`nAttachment uploaded successfully!" -ForegroundColor Green
    Write-Host "Response:" -ForegroundColor Cyan
    $result | ConvertTo-Json -Depth 10 | Write-Host -ForegroundColor Gray
    
    return $result
}
catch {
    Write-Error "Failed to upload attachment: $($_.Exception.Message)"
    throw
}

