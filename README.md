# OutlookTools

A PowerShell toolkit for managing and extracting data from Microsoft Outlook mailboxes.

## Key Features

### Functions
- **Initialize-Outlook**: Creates and initializes the Outlook COM object
- **List-OutlookFolders**: Recursively lists all folders in the mailbox with item counts
- **Extract-OutlookEmails**: Extracts names and email addresses from a specific folder

### Parameters
- `-ListFolders`: Lists all folders and saves to a CSV file
- `-ExtractEmails`: Extracts emails from a specific folder and saves to a CSV file
- `-FolderPath`: Specifies the folder path for email extraction
- `-MailboxName`: Specifies the mailbox name (optional)

## Usage

1. Save the script as `OutlookScript.ps1`
2. Run using the following command patterns:

### List folders in mailbox:
```powershell
.\OutlookScript.ps1 -GetMailboxFolders -MailboxName "Mailbox - John Doe" -OutputFolder "C:\Reports"
```

### Extract email recipient addresses and names:
```powershell
.\OutlookScript.ps1 -GetEmailsFromFolder -FolderPath "Inbox\MyFolder" -MailboxName "Mailbox - John Doe" -OutputFolder "D:\Exports"
```

### Use default output to Desktop:
```powershell
.\OutlookScript.ps1 -GetEmailsFromFolder -FolderPath "Inbox\MyFolder" -MailboxName "Mailbox - John Doe"
```

## Output Files

### OutlookFolderList.csv
Contains folder listing with the following columns:
- **Mailbox**: The mailbox name (e.g., artclown@somedomainsomewhere.com)
- **FolderPath**: The path to the folder within the mailbox (e.g., \Inbox\Blah)
- **ItemCount**: The number of items in the folder

### ExtractedEmails.csv
Contains extracted email addresses and names.

### File Handling
If an output file already exists, the script will:
1. Prompt for overwrite confirmation
2. If declined, create a new file with timestamp (e.g., OutlookFolderList_20250108123000.csv)

## Outlook Integration

The script uses the Outlook COM Object and can work with both running and closed instances of Outlook:

- If Outlook is running, the script connects to the existing instance
- If Outlook is closed, the script launches it automatically in the background

### Important Considerations

1. **During Operation**
   - You can continue working with Outlook while the script runs
   - Avoid modifying emails in the target folder during script execution

2. **Performance**
   - Large mailboxes or folders may temporarily impact Outlook performance
   - Consider closing resource-intensive operations before running the script

3. **Security**
   - Organization settings may trigger security prompts when accessing mailbox data
   - These prompts are normal security measures

4. **Best Practices**
   - Close other heavy tasks or large folder operations in Outlook
   - This minimizes potential performance or locking issues

## PowerShell Help
The script includes comprehensive PowerShell help comments for easy reference and maintenance.
