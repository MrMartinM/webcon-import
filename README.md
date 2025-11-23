# Webcon PowerShell Automation Tools

A collection of PowerShell tools for automating Webcon workflow operations via the REST API.

## Overview

This repository contains multiple standalone tools for interacting with Webcon workflows:

1. **ImportWorkflows** - Import data from Excel files into Webcon workflows
2. **UploadAttachments** - Upload files as attachments to Webcon elements

Each tool is self-contained with its own modules, configuration, and documentation.

## Available Tools

### 1. ImportWorkflows

Automatically import data from Excel files into Webcon workflows. Supports field mapping, item lists, progress tracking, and automatic resume capability.

**Features:**
- Excel-based configuration (no code changes needed)
- Automatic resume (skips already imported rows)
- Progress window with real-time status
- Automatic retry on network errors
- Status tracking in CSV file
- Support for item lists

**Documentation:** See [ImportWorkflows/README.md](ImportWorkflows/README.md)

### 2. UploadAttachments

Upload files from your local disk to Webcon elements as attachments via the REST API.

**Features:**
- Automatic authentication with OAuth2
- Base64 encoding of file content
- Automatic retry on network errors
- Support for large files
- UTF-8 encoding support

**Documentation:** See [UploadAttachments/README.md](UploadAttachments/README.md)

## Common Prerequisites

- PowerShell 5.1 or later
- Webcon instance with REST API access
- OAuth2 client credentials (ClientId and ClientSecret)

## Project Structure

```
webcon-import/
├── README.md (this file)
├── ImportWorkflows/
│   ├── README.md
│   ├── Start-WebconWorkflows.ps1
│   ├── Config.json.example
│   ├── TECHNICAL.md
│   ├── Modules/
│   │   ├── ExcelReader.psm1
│   │   ├── ProgressWindow.psm1
│   │   ├── StatusTracker.psm1
│   │   └── WebconAPI.psm1
│   └── SQL/
│       ├── mme_GetFieldDefinitions.sql
│       └── mme_GetFieldDetailDefinitions.sql
└── UploadAttachments/
    ├── README.md
    ├── Add-WebconAttachment.ps1
    ├── Config.json.example
    └── Modules/
        └── WebconAPI.psm1
```

## Quick Start

### ImportWorkflows

1. Navigate to the ImportWorkflows folder:
   ```powershell
   cd ImportWorkflows
   ```

2. Copy and configure `Config.json.example`:
   ```powershell
   Copy-Item Config.json.example Config.json
   # Edit Config.json with your settings
   ```

3. Run the script:
   ```powershell
   .\Start-WebconWorkflows.ps1
   ```

See [ImportWorkflows/README.md](ImportWorkflows/README.md) for detailed instructions.

### UploadAttachments

1. Navigate to the UploadAttachments folder:
   ```powershell
   cd UploadAttachments
   ```

2. Copy and configure `Config.json.example`:
   ```powershell
   Copy-Item Config.json.example Config.json
   # Edit Config.json with your settings
   ```

3. Run the script:
   ```powershell
   .\Add-WebconAttachment.ps1 -FilePath "C:\path\to\file.pdf" -ElementId 123
   ```

See [UploadAttachments/README.md](UploadAttachments/README.md) for detailed instructions.

## Security

Both tools support secure credential management:

- **Recommended**: Set `ClientSecret` as an environment variable:
  ```powershell
  [System.Environment]::SetEnvironmentVariable("WEBCON_CLIENT_SECRET", "your-secret", "User")
  ```
  Then leave `ClientSecret` empty in `Config.json`.

- **Alternative**: Provide `ClientSecret` directly in `Config.json` (less secure).

## Getting Help

- **ImportWorkflows**: See [ImportWorkflows/README.md](ImportWorkflows/README.md) and [ImportWorkflows/TECHNICAL.md](ImportWorkflows/TECHNICAL.md)
- **UploadAttachments**: See [UploadAttachments/README.md](UploadAttachments/README.md)

## License

This project is provided as-is for use with Webcon BPS.
