# WebconAPI.psm1
# Module for Webcon API authentication and attachment operations

function Get-WebconToken {
    <#
    .SYNOPSIS
    Authenticates with Webcon API and returns access token
    
    .PARAMETER BaseUrl
    Base URL of the Webcon instance
    
    .PARAMETER ClientId
    OAuth2 client ID
    
    .PARAMETER ClientSecret
    OAuth2 client secret
    
    .EXAMPLE
    $token = Get-WebconToken -BaseUrl "https://test-webcon.dragonmaritime.si" -ClientId "xxx" -ClientSecret "xxx"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$ClientId,
        
        [Parameter(Mandatory=$true)]
        [string]$ClientSecret
    )
    
    $tokenUrl = "$BaseUrl/api/oauth2/token"
    
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        
        if ($response.access_token) {
            return $response.access_token
        }
        else {
            throw "No access token received in response"
        }
    }
    catch {
        Write-Error "Failed to get access token: $($_.Exception.Message)"
        throw
    }
}

function Add-WebconAttachment {
    <#
    .SYNOPSIS
    Adds an attachment to a Webcon element
    
    .PARAMETER BaseUrl
    Base URL of the Webcon instance
    
    .PARAMETER AccessToken
    Bearer token for authentication
    
    .PARAMETER DatabaseId
    Database ID (e.g., 9)
    
    .PARAMETER ElementId
    Element ID to attach file to
    
    .PARAMETER Name
    Attachment name
    
    .PARAMETER Description
    Attachment description (optional)
    
    .PARAMETER Group
    Attachment group (optional)
    
    .PARAMETER Content
    Base64 encoded file content
    
    .EXAMPLE
    Add-WebconAttachment -BaseUrl "https://..." -AccessToken $token -DatabaseId 9 -ElementId 123 -Name "document.pdf" -Content $base64Content
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$AccessToken,
        
        [Parameter(Mandatory=$true)]
        [string]$DatabaseId,
        
        [Parameter(Mandatory=$true)]
        [int]$ElementId,
        
        [Parameter(Mandatory=$true)]
        [string]$Name,
        
        [Parameter(Mandatory=$false)]
        [string]$Description = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Group = "",
        
        [Parameter(Mandatory=$true)]
        [string]$Content
    )
    
    $apiUrl = "$BaseUrl/api/data/v6.0/db/$DatabaseId/elements/$ElementId/attachments"
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Accept"        = "text/plain"
        "Content-Type"  = "application/json"
    }
    
    # Build body
    $body = [ordered]@{
        name        = $Name
        description = $Description
        group       = $Group
        content     = $Content
    }
    
    try {
        # Convert to JSON with proper UTF-8 encoding
        $jsonBody = $body | ConvertTo-Json -Depth 10
        
        # Ensure the JSON string is properly UTF-8 encoded
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        $jsonBytes = $utf8NoBom.GetBytes($jsonBody)
        $jsonBodyUtf8 = $utf8NoBom.GetString($jsonBytes)
        
        Write-Host "Uploading attachment: $Name" -ForegroundColor Cyan
        
        # Use Invoke-WebRequest with UTF-8 encoded bytes
        try {
            $webRequest = Invoke-WebRequest -Uri $apiUrl -Method Post -Headers $headers -Body $jsonBytes -ContentType "application/json; charset=utf-8" -UseBasicParsing
            $response = $webRequest.Content | ConvertFrom-Json
            return $response
        }
        catch {
            $errorResponse = $_.Exception.Response
            $errorMessage = "Failed to add attachment: $($_.Exception.Message)"
            $responseBody = $null
            
            # Try to read the response body
            if ($errorResponse) {
                try {
                    $responseStream = $errorResponse.GetResponseStream()
                    if ($responseStream -and $responseStream.CanRead) {
                        # Ensure stream is at the beginning
                        $responseStream.Position = 0
                        
                        # Try to determine encoding from Content-Type header
                        $encoding = [System.Text.Encoding]::UTF8
                        if ($errorResponse.ContentType) {
                            if ($errorResponse.ContentType -match "charset=([^;]+)") {
                                try {
                                    $encoding = [System.Text.Encoding]::GetEncoding($matches[1])
                                } catch {
                                    $encoding = [System.Text.Encoding]::UTF8
                                }
                            }
                        }
                        
                        $reader = New-Object System.IO.StreamReader($responseStream, $encoding)
                        $responseBody = $reader.ReadToEnd()
                        $reader.Close()
                        $responseStream.Close()
                    }
                }
                catch {
                    Write-Host "Could not read response stream: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # Also try to get response from ErrorRecord/ErrorDetails
            if (-not $responseBody) {
                if ($_.ErrorDetails) {
                    $responseBody = $_.ErrorDetails.Message
                }
                elseif ($_.Exception.Response) {
                    # Try alternative method
                    try {
                        $memStream = New-Object System.IO.MemoryStream
                        $_.Exception.Response.GetResponseStream().CopyTo($memStream)
                        $memStream.Position = 0
                        $reader = New-Object System.IO.StreamReader($memStream)
                        $responseBody = $reader.ReadToEnd()
                        $reader.Close()
                        $memStream.Close()
                    } catch {
                        # Ignore
                    }
                }
            }
            
            if ($responseBody) {
                Write-Host "Response body:" -ForegroundColor Red
                Write-Host $responseBody -ForegroundColor Yellow
                
                # Try to parse as JSON for better error message
                try {
                    $errorObj = $responseBody | ConvertFrom-Json
                    
                    # Handle ValidationError structure
                    if ($errorObj.type -and $errorObj.description) {
                        $errorMessage += " - Type: $($errorObj.type), Description: $($errorObj.description)"
                        if ($errorObj.errorGuid) {
                            $errorMessage += ", ErrorGuid: $($errorObj.errorGuid)"
                        }
                    }
                    # Handle other error formats
                    elseif ($errorObj.message) {
                        $errorMessage += " - $($errorObj.message)"
                    } elseif ($errorObj.error) {
                        $errorMessage += " - $($errorObj.error)"
                    } elseif ($errorObj.Message) {
                        $errorMessage += " - $($errorObj.Message)"
                    } elseif ($errorObj.description) {
                        $errorMessage += " - $($errorObj.description)"
                    }
                    else {
                        # If we have a JSON object but no recognized fields, show the full object
                        $errorMessage += " - Response: $responseBody"
                    }
                } catch {
                    # If not JSON, use raw response
                    $errorMessage += " - Response: $responseBody"
                }
            }
            
            Write-Host $errorMessage -ForegroundColor Red
            Write-Error $errorMessage
            throw
        }
    }
    catch {
        # Re-throw if already handled above
        throw
    }
}

function Add-WebconAttachmentWithRetry {
    <#
    .SYNOPSIS
    Adds an attachment to a Webcon element with retry logic for transient errors
    
    .PARAMETER BaseUrl
    Base URL of the Webcon instance
    
    .PARAMETER AccessToken
    Bearer token for authentication
    
    .PARAMETER DatabaseId
    Database ID (e.g., 9)
    
    .PARAMETER ElementId
    Element ID to attach file to
    
    .PARAMETER Name
    Attachment name
    
    .PARAMETER Description
    Attachment description (optional)
    
    .PARAMETER Group
    Attachment group (optional)
    
    .PARAMETER Content
    Base64 encoded file content
    
    .PARAMETER MaxRetries
    Maximum number of retry attempts (default: 3)
    
    .PARAMETER RetryDelayBase
    Base delay in seconds for exponential backoff (default: 2)
    
    .EXAMPLE
    Add-WebconAttachmentWithRetry -BaseUrl "https://..." -AccessToken $token -DatabaseId 9 -ElementId 123 -Name "document.pdf" -Content $base64Content
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$AccessToken,
        
        [Parameter(Mandatory=$true)]
        [string]$DatabaseId,
        
        [Parameter(Mandatory=$true)]
        [int]$ElementId,
        
        [Parameter(Mandatory=$true)]
        [string]$Name,
        
        [Parameter(Mandatory=$false)]
        [string]$Description = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Group = "",
        
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$RetryDelayBase = 2
    )
    
    $attempt = 0
    $lastException = $null
    
    while ($attempt -le $MaxRetries) {
        try {
            $result = Add-WebconAttachment -BaseUrl $BaseUrl `
                                          -AccessToken $AccessToken `
                                          -DatabaseId $DatabaseId `
                                          -ElementId $ElementId `
                                          -Name $Name `
                                          -Description $Description `
                                          -Group $Group `
                                          -Content $Content
            
            if ($attempt -gt 0) {
                Write-Host "Attachment uploaded successfully on attempt $($attempt + 1)" -ForegroundColor Green
            }
            
            return $result
        }
        catch {
            $lastException = $_
            $shouldRetry = $false
            $isRetryableError = $false
            
            # Check if this is a retryable error
            if ($_.Exception -is [System.Net.WebException]) {
                $webException = $_.Exception
                
                # Check for timeout
                if ($webException.Status -eq [System.Net.WebExceptionStatus]::Timeout -or
                    $webException.Status -eq [System.Net.WebExceptionStatus]::ConnectFailure -or
                    $webException.Status -eq [System.Net.WebExceptionStatus]::ReceiveFailure) {
                    $isRetryableError = $true
                }
                
                # Check HTTP status code
                if ($webException.Response) {
                    $httpResponse = $webException.Response
                    if ($httpResponse -is [System.Net.HttpWebResponse]) {
                        $statusCode = [int]$httpResponse.StatusCode
                        
                        # Retry on server errors (5xx) except 501 (Not Implemented)
                        if ($statusCode -ge 500 -and $statusCode -ne 501) {
                            $isRetryableError = $true
                        }
                        # Don't retry on client errors (4xx) - these are permanent failures
                        elseif ($statusCode -ge 400 -and $statusCode -lt 500) {
                            $isRetryableError = $false
                            $shouldRetry = $false
                        }
                    }
                }
            }
            elseif ($_.Exception -is [System.TimeoutException]) {
                $isRetryableError = $true
            }
            
            # Determine if we should retry
            if ($isRetryableError -and $attempt -lt $MaxRetries) {
                $shouldRetry = $true
            }
            
            if ($shouldRetry) {
                $attempt++
                $delay = $RetryDelayBase * [Math]::Pow(2, $attempt - 1)
                
                Write-Warning "Attempt $attempt failed: $($_.Exception.Message). Retrying in $delay seconds..."
                Start-Sleep -Seconds $delay
            }
            else {
                # Don't retry - either max retries reached or permanent error
                if ($attempt -ge $MaxRetries) {
                    Write-Error "Failed after $MaxRetries retry attempts. Last error: $($_.Exception.Message)"
                }
                else {
                    Write-Error "Permanent error (not retryable): $($_.Exception.Message)"
                }
                throw $lastException
            }
        }
    }
    
    # Should never reach here, but just in case
    throw $lastException
}

Export-ModuleMember -Function Get-WebconToken, Add-WebconAttachment, Add-WebconAttachmentWithRetry

