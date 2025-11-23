# Webcon Attachment Upload Tool

Upload files as attachments to Webcon workflow elements via the REST API.

## Overview

This tool allows you to upload files from your local disk to Webcon elements as attachments. It handles authentication, file encoding (base64), and includes automatic retry logic for transient errors.

## Prerequisites

- PowerShell 5.1 or later
- Webcon instance with REST API access
- OAuth2 client credentials (ClientId and ClientSecret)

## Installation

No additional modules required - uses only built-in PowerShell cmdlets.

## Configuration

1. Copy `Config.json.example` to `Config.json`:
   ```powershell
   Copy-Item Config.json.example Config.json
   ```

2. Edit `Config.json`:
   ```json
   {
     "Webcon": {
       "BaseUrl": "https://your-webcon-instance.com",
       "ClientId": "your-client-id-here",
       "ClientSecret": "",
       "DatabaseId": "9"
     },
     "Retry": {
       "MaxRetries": 3,
       "RetryDelayBase": 2
     }
   }
   ```

3. **Security**: Leave `ClientSecret` empty and set it as an environment variable:
   ```powershell
   [System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-secret", "User")
   ```

## Usage

### Basic Usage

Upload a file to a Webcon element:

```powershell
cd UploadAttachments
.\Add-WebconAttachment.ps1 -FilePath "C:\path\to\file.pdf" -ElementId 123
```

### With Optional Parameters

```powershell
.\Add-WebconAttachment.ps1 `
    -FilePath "C:\documents\report.pdf" `
    -ElementId 123 `
    -Name "Monthly Report" `
    -Description "Monthly sales report for Q1" `
    -Group "Reports" `
    -ConfigPath ".\Config.json"
```

### Parameters

- **FilePath** (Mandatory): Path to the file on disk
- **ElementId** (Mandatory): Webcon element ID to attach the file to
- **Name** (Optional): Attachment name (defaults to filename if not provided)
- **Description** (Optional): Attachment description (defaults to empty string)
- **Group** (Optional): Attachment group (defaults to empty string)
- **ConfigPath** (Optional): Path to Config.json (defaults to `.\Config.json`)

## API Endpoint

The script uses the Webcon REST API endpoint:

```
POST /api/data/v6.0/db/{dbid}/elements/{id}/attachments
```

**Request Body:**
```json
{
  "name": "string",
  "description": "string",
  "group": "string",
  "content": "base64-encoded-file-content"
}
```

**Response:**
Returns a JSON object with attachment details.

## Features

- ✅ Automatic authentication with OAuth2
- ✅ Base64 encoding of file content
- ✅ Automatic retry on network errors (configurable)
- ✅ Detailed error messages
- ✅ Support for large files
- ✅ UTF-8 encoding support for special characters

## Troubleshooting

**"ClientSecret not found"**
- Set environment variable: `[System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-secret", "User")`
- Or provide it in `Config.json` (less secure)

**"File not found"**
- Use absolute path for `-FilePath`
- Ensure file is not locked by another process

**"Failed to add attachment"**
- Verify ElementId is correct and element exists
- Check that the OAuth2 client has permissions to add attachments
- Verify DatabaseId matches your Webcon database

**"Failed to authenticate"**
- Verify BaseUrl, ClientId, and ClientSecret are correct
- Check network connectivity to Webcon instance

**Large file uploads**
- The script handles large files by reading them in binary format and converting to base64
- For very large files (>100MB), consider network timeout settings

## Retry Logic

The script automatically retries on:
- Network timeouts
- Connection failures
- Server errors (5xx, except 501)

It does NOT retry on:
- Client errors (4xx) - these are permanent failures
- Authentication errors

Retry behavior can be configured in `Config.json`:
- `MaxRetries`: Maximum number of retry attempts (default: 3)
- `RetryDelayBase`: Base delay in seconds for exponential backoff (default: 2)

## Examples

### Upload a PDF document

```powershell
.\Add-WebconAttachment.ps1 -FilePath "C:\reports\report.pdf" -ElementId 456
```

### Upload an image with description

```powershell
.\Add-WebconAttachment.ps1 `
    -FilePath "C:\images\screenshot.png" `
    -ElementId 789 `
    -Name "Screenshot" `
    -Description "Application screenshot"
```

### Upload to a specific group

```powershell
.\Add-WebconAttachment.ps1 `
    -FilePath "C:\data\export.csv" `
    -ElementId 123 `
    -Group "Data Exports"
```

