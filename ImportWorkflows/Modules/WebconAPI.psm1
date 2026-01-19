# WebconAPI.psm1
# Module for Webcon API authentication and workflow operations

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

function Start-WebconWorkflow {
    <#
    .SYNOPSIS
    Starts a workflow in Webcon by creating a new element
    
    .PARAMETER BaseUrl
    Base URL of the Webcon instance
    
    .PARAMETER AccessToken
    Bearer token for authentication
    
    .PARAMETER DatabaseId
    Database ID (e.g., 9)
    
    .PARAMETER WorkflowGuid
    GUID of the workflow to start
    
    .PARAMETER FormTypeGuid
    GUID of the form type
    
    .PARAMETER FormFields
    Array of form field objects
    
    .PARAMETER ItemLists
    Array of item list objects (optional)
    
    .PARAMETER Path
    Path parameter (default: "default")
    
    .PARAMETER Mode
    Mode parameter (default: "standard")
    
    .PARAMETER BusinessEntityGuid
    Optional business entity GUID to set on the element
    
    .EXAMPLE
    Start-WebconWorkflow -BaseUrl "https://..." -AccessToken $token -DatabaseId 9 -WorkflowGuid "..." -FormTypeGuid "..." -FormFields $fields
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
        [string]$WorkflowGuid,
        
        [Parameter(Mandatory=$true)]
        [string]$FormTypeGuid,
        
        [Parameter(Mandatory=$false)]
        [array]$FormFields = @(),
        
        [Parameter(Mandatory=$false)]
        [array]$ItemLists = @(),
        
        [Parameter(Mandatory=$false)]
        [string]$Path = "default",
        
        [Parameter(Mandatory=$false)]
        [string]$Mode = "standard",
        
        [Parameter(Mandatory=$false)]
        [string]$BusinessEntityGuid
    )
    
    $apiUrl = "$BaseUrl/api/data/v6.0/db/$DatabaseId/elements"
    $uri = "$apiUrl" + "?path=$Path&mode=$Mode"
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Accept"        = "text/plain"
        "Content-Type"  = "application/json"
    }
    
    # Build body with specific order: workflow, formType, formFields
    $body = [ordered]@{
        workflow = [ordered]@{
            guid = $WorkflowGuid
        }
        formType = [ordered]@{
            guid = $FormTypeGuid
        }
        formFields = $FormFields
    }
    
    # Add businessEntity if provided
    if ($BusinessEntityGuid -and $BusinessEntityGuid.Trim() -ne "") {
        $body.businessEntity = [ordered]@{
            guid = $BusinessEntityGuid
        }
    }
    
    # Add itemLists if provided
    if ($ItemLists.Count -gt 0) {
        $body.itemLists = $ItemLists
    }
    
    try {
        # Convert to JSON with proper UTF-8 encoding
        $jsonBody = $body | ConvertTo-Json -Depth 10
        
        # Ensure the JSON string is properly UTF-8 encoded
        # Convert to UTF-8 bytes to handle special characters correctly
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        $jsonBytes = $utf8NoBom.GetBytes($jsonBody)
        $jsonBodyUtf8 = $utf8NoBom.GetString($jsonBytes)
        
        Write-Host "Request JSON:" -ForegroundColor Cyan
        Write-Host $jsonBodyUtf8 -ForegroundColor Gray
        
        # Use Invoke-WebRequest with UTF-8 encoded bytes to ensure proper encoding
        # This prevents encoding issues with special characters like ø (byte F8)
        try {
            $webRequest = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -Body $jsonBytes -ContentType "application/json; charset=utf-8" -UseBasicParsing
            $response = $webRequest.Content | ConvertFrom-Json
            return $response
        }
        catch {
            $errorResponse = $_.Exception.Response
            $errorMessage = "Failed to start workflow: $($_.Exception.Message)"
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

function Start-WebconWorkflowWithRetry {
    <#
    .SYNOPSIS
    Starts a workflow in Webcon with retry logic for transient errors
    
    .PARAMETER BaseUrl
    Base URL of the Webcon instance
    
    .PARAMETER AccessToken
    Bearer token for authentication
    
    .PARAMETER DatabaseId
    Database ID (e.g., 9)
    
    .PARAMETER WorkflowGuid
    GUID of the workflow to start
    
    .PARAMETER FormTypeGuid
    GUID of the form type
    
    .PARAMETER FormFields
    Array of form field objects
    
    .PARAMETER ItemLists
    Array of item list objects (optional)
    
    .PARAMETER Path
    Path parameter (default: "default")
    
    .PARAMETER Mode
    Mode parameter (default: "standard")
    
    .PARAMETER BusinessEntityGuid
    Optional business entity GUID to set on the element
    
    .PARAMETER MaxRetries
    Maximum number of retry attempts (default: 3)
    
    .PARAMETER RetryDelayBase
    Base delay in seconds for exponential backoff (default: 2)
    
    .EXAMPLE
    Start-WebconWorkflowWithRetry -BaseUrl "https://..." -AccessToken $token -DatabaseId 9 -WorkflowGuid "..." -FormTypeGuid "..." -FormFields $fields
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
        [string]$WorkflowGuid,
        
        [Parameter(Mandatory=$true)]
        [string]$FormTypeGuid,
        
        [Parameter(Mandatory=$false)]
        [array]$FormFields = @(),
        
        [Parameter(Mandatory=$false)]
        [array]$ItemLists = @(),
        
        [Parameter(Mandatory=$false)]
        [string]$Path = "default",
        
        [Parameter(Mandatory=$false)]
        [string]$Mode = "standard",
        
        [Parameter(Mandatory=$false)]
        [string]$BusinessEntityGuid,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$RetryDelayBase = 2
    )
    
    $attempt = 0
    $lastException = $null
    
    while ($attempt -le $MaxRetries) {
        try {
            $result = Start-WebconWorkflow -BaseUrl $BaseUrl `
                                           -AccessToken $AccessToken `
                                           -DatabaseId $DatabaseId `
                                           -WorkflowGuid $WorkflowGuid `
                                           -FormTypeGuid $FormTypeGuid `
                                           -FormFields $FormFields `
                                           -ItemLists $ItemLists `
                                           -Path $Path `
                                           -Mode $Mode `
                                           -BusinessEntityGuid $BusinessEntityGuid
            
            if ($attempt -gt 0) {
                Write-Host "Workflow started successfully on attempt $($attempt + 1)" -ForegroundColor Green
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

Export-ModuleMember -Function Get-WebconToken, Start-WebconWorkflow, Start-WebconWorkflowWithRetry

