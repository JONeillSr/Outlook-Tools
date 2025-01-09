# Send Email From Outlook

A PowerShell script that automates sending personalized emails through Microsoft Outlook using Word templates and CSV recipient lists.

## Features

- Sends personalized emails using Microsoft Outlook
- Uses Word documents as email templates
- Supports HTML formatting and embedded images
- Personalizes content using placeholders (e.g., [GivenName])
- Processes multiple recipients from a CSV file
- Provides detailed logging of all operations
- Maintains original Word document formatting
- Handles both new and existing Outlook/Word instances

## Prerequisites

- Microsoft Office 2016 or later
- Microsoft Outlook (installed and configured with valid email account)
- Microsoft Word
- PowerShell 5.1 or later
- Appropriate permissions to run PowerShell scripts

## Installation

1. Save `sendemailfromoutlook.ps1` to your desired location
2. If needed, unblock the file:
```powershell
Unblock-File -Path .\sendemailfromoutlook.ps1
```
3. Set appropriate PowerShell execution policy:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Usage

### Basic Syntax
```powershell
.\sendemailfromoutlook.ps1 -InputTemplate "path\to\template.docx" -EmailSubject "Your Subject" -InputCSV "path\to\recipients.csv"
```

### Parameters

- `-InputTemplate`: Path to Word document (.docx) template
- `-EmailSubject`: Subject line for the emails
- `-InputCSV`: Path to CSV file containing recipient information
- `-Verbose`: (Optional) Enable detailed progress output

### Examples

Basic usage:
```powershell
.\sendemailfromoutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv"
```

With verbose logging:
```powershell
.\sendemailfromoutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -Verbose
```

## CSV Format

The input CSV file must contain the following columns:
```csv
Name,GivenName,Surname,Email
Art Clown,Art,Clown,arttheclown@somedomainsomewhere.com
```

Required columns:
- `Name`: Full name of recipient
- `GivenName`: First name of recipient
- `Surname`: Last name of recipient
- `Email`: Email address of recipient

## Template Formatting

- Create your email template as a regular Word document (.docx)
- Use [GivenName] as a placeholder to be replaced with the recipient's first name
- Formatting from the Word document (including images) will be preserved in the email

## Logging

The script automatically creates detailed log files:

- Location: `%LOCALAPPDATA%\PowerShell\logs`
- Filename format: `SendEmailFromOutlook_YYYY-MM-DD_HH-mm-ss.log`
- Content: Timestamps and detailed operation information
- Each operation and error is logged for troubleshooting

## Error Handling

The script includes comprehensive error handling:
- Validates input files before processing
- Checks CSV format and required columns
- Verifies Outlook configuration
- Handles COM object cleanup
- Logs all errors with detailed information

## Author

- Author: John A. O'Neill Sr.
- Version: 1.1
- Last Updated: 01/08/2025

## Notes

1. The script uses COM objects to interact with Outlook and Word
2. Outlook security prompts may appear based on your organization's settings
3. A small delay (500ms) is added between emails to prevent throttling
4. The script can work with both running and closed instances of Outlook/Word
5. All COM objects are properly cleaned up after use

## Support

If you encounter any issues:
1. Check the log files for detailed error information
2. Verify all prerequisites are met
3. Ensure input files meet the required format
4. Run with -Verbose flag for additional debugging information