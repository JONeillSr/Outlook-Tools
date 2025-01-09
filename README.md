# Outlook PowerShell Tools

A collection of PowerShell scripts for automating Microsoft Outlook tasks, mailbox management, and email communications.

## Available Tools

### 1. Outlook Tools
Located in `/outlooktools`

A toolkit for managing and analyzing Outlook mailbox data:
- Recursive folder listing with item counts
- Email address extraction from folders
- Detailed folder analysis
- CSV export of mailbox structure
- Support for multiple mailboxes

[View Outlook Tools Documentation](OutlookTools/readme.md)

### 2. Send Email from Outlook
Located in `/sendemailfromoutlook`

An automated email sending solution that combines Word templates with CSV recipient lists:
- Personalized email sending using Word templates
- Mail merge functionality with CSV data
- HTML formatting preservation
- Support for embedded images
- Comprehensive logging system
- Bulk email processing with throttling protection
- Sends emails individually for a more personal interaction and to prevent recipients from harvesting the addresses of other recipients

[View Send Email Documentation](SendEmailFromOutlook/readme.md)

## Getting Started

### Prerequisites
- Windows PowerShell 5.1 or later
- Microsoft Outlook (Desktop version)
- Microsoft Word (Required for sendemailfromoutlook)
- Appropriate permissions to run PowerShell scripts in your environment
- Valid Outlook email account configuration

### Installation
1. Clone this repository:
```powershell
git clone https://github.com/JONeillSr/outlook-tools.git
```

2. Navigate to the desired tool directory:
```powershell
cd outlook-tools/outlooktools
# or
cd outlook-tools/sendemailfromoutlook
```

3. Unblock the PowerShell scripts if needed:
```powershell
Unblock-File -Path .\outlooktools.ps1
# or
Unblock-File -Path .\sendemailfromoutlook.ps1
```

4. Set appropriate PowerShell execution policy:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Common Features Across Tools

### Outlook Integration
Both tools:
- Work with running or closed instances of Outlook
- Use COM object interaction
- Handle Outlook security prompts
- Provide proper cleanup of COM objects

### Data Export
Both tools export data to CSV format:
- outlooktools: Mailbox structure and email extraction
- sendemailfromoutlook: Logging and operation records

### Logging
- outlooktools: Logs folder operations and email extraction
- sendemailfromoutlook: Detailed logging in %LOCALAPPDATA%\PowerShell\logs

## Security Considerations

These scripts interact with Outlook using COM objects. Depending on your organization's security settings, you may need to:
- Unblock downloaded PS1 files
- Set appropriate PowerShell execution policies
- Handle Outlook security prompts
- Ensure proper email account permissions

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Guidelines
1. Maintain consistent formatting with the existing codebase
2. Update documentation for any modified functions
3. Test your changes thoroughly
4. Follow PowerShell best practices
5. Include appropriate error handling
6. Add logging for new functionality

## Authors
- Original Author: John A. O'Neill Sr.
- Last Updated: January 8, 2025

## License

[MIT License](LICENSE)

## Project Structure
```
outlook-tools/
│
├── outlooktools/
│   ├── outlooktools.ps1
│   └── README.md
│
└── sendemailfromoutlook/
    ├── sendemailfromoutlook.ps1
    └── README.md
    └── testinput.csv
    └── Sample.docx
```

## Support

If you encounter any issues or have questions:
1. Check the specific tool's README for common issues
2. Review the log files for error details
3. Verify prerequisites are met
4. Run scripts with -Verbose flag for additional debugging information
5. Open an issue in the GitHub repository
6. Provide relevant details about your environment and the error encountered
