# Outlook Tools

A PowerShell toolkit for comprehensive Outlook mailbox management and analysis, providing detailed folder listing and email data extraction capabilities.

## Features

- **Mailbox Folder Management**:
  - Recursive folder listing with item counts
  - Support for multiple mailboxes
  - Detailed folder path reporting
  - CSV export of mailbox structure

- **Email Data Extraction**:
  - Extract names and email addresses from specific folders
  - Bulk processing capabilities
  - Customizable output formats
  - Detailed extraction logging

- **Folder Analysis**:
  - Item count reporting
  - Path structure analysis
  - Custom folder filtering
  - Multiple mailbox support

## Prerequisites

- Windows PowerShell 5.1 or later
- Microsoft Outlook (Desktop version)
- Appropriate permissions to run PowerShell scripts
- Valid Outlook email account configuration

## Installation

1. Save `outlooktools.ps1` to your desired location
2. If needed, unblock the file:
```powershell
Unblock-File -Path .\outlooktools.ps1
```
3. Set appropriate PowerShell execution policy:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Usage

### List Folders in Mailbox
```powershell
.\outlooktools.ps1 -GetMailboxFolders -MailboxName "Mailbox - John Doe" -OutputFolder "C:\Reports"
```

### Extract Email Addresses
```powershell
.\outlooktools.ps1 -GetEmailsFromFolder -FolderPath "Inbox\MyFolder" -MailboxName "Mailbox - John Doe" -OutputFolder "D:\Exports"
```

### Use Default Output Location (Desktop)
```powershell
.\outlooktools.ps1 -GetEmailsFromFolder -FolderPath "Inbox\MyFolder" -MailboxName "Mailbox - John Doe"
```

## Parameters

- `-GetMailboxFolders`: Lists all folders and saves to CSV
- `-GetEmailsFromFolder`: Extracts emails from specified folder
- `-FolderPath`: Specifies the folder path for operations
- `-MailboxName`: Specifies the mailbox name (optional)
- `-OutputFolder`: Destination for CSV files (optional, defaults to Desktop)

## Output Files

### OutlookFolderList.csv
Contains folder structure information:
- `Mailbox`: The mailbox name (e.g., artclown@somedomainsomewhere.com)
- `FolderPath`: Full path to the folder (e.g., \Inbox\Subfolder)
- `ItemCount`: Number of items in the folder

### ExtractedEmails.csv
Contains email data from specified folders.

### File Handling
When output files already exist:
1. Script prompts for overwrite confirmation
2. If declined, creates new file with timestamp (e.g., OutlookFolderList_20250108123000.csv)

## Outlook Integration

The script works seamlessly with Outlook:
- Compatible with running or closed Outlook instances
- Uses Outlook COM object for reliable interaction
- Proper cleanup of COM objects after use
- Handles Outlook security prompts appropriately

### Working with Running Outlook
- Safe to run while Outlook is open
- Avoid modifying target folders during script execution
- May experience temporary performance impact with large folders

## Error Handling

The script includes comprehensive error handling:
- Validates input parameters
- Checks folder paths
- Verifies mailbox access
- Handles COM object errors
- Provides detailed error messages

## Security Notes

1. **COM Object Access**:
   - Script requires appropriate permissions
   - May trigger Outlook security prompts
   - Respects Outlook security settings

2. **File System Access**:
   - Requires write permissions for output location
   - Creates directories if needed
   - Handles file access conflicts

3. **Mailbox Access**:
   - Requires appropriate mailbox permissions
   - Works with default and additional mailboxes
   - Respects mailbox access restrictions

## Performance Considerations

- Large mailboxes may require extended processing time
- Consider running during off-peak hours for large operations
- Script includes built-in handling for large folder structures
- Memory usage scales with folder size and depth

## Support

If you encounter issues:
1. Verify prerequisites are met
2. Check your Outlook connection and permissions
3. Run with detailed logging for troubleshooting:
```powershell
.\outlooktools.ps1 -GetMailboxFolders -MailboxName "Mailbox - John Doe" -Verbose
```
4. Review error messages and output
5. Check for adequate permissions
6. Verify Outlook configuration

## Notes

1. The script uses COM objects to interact with Outlook
2. Outlook security prompts may appear based on your settings
3. Large mailboxes may require additional processing time
4. All COM objects are properly cleaned up after use

## Best Practices

1. Regular backups before bulk operations
2. Test with small folder sets first
3. Use verbose mode for detailed logging
4. Monitor system resources during large operations
5. Follow proper error handling procedures

## Author

- Author: John A. O'Neill Sr.
- Last Updated: January 8, 2025