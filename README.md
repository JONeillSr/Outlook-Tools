# Outlook PowerShell Tools

A collection of PowerShell scripts for automating Microsoft Outlook tasks and email management.

## Available Tools

### 1. Outlook Tools
Located in `/outlooktools`

A comprehensive toolkit for managing Outlook folders and extracting email data. Features include:
- Recursive folder listing with item counts
- Email address extraction
- Detailed folder analysis
- CSV export capabilities

[View Outlook Tools Documentation](./outlooktools/README.md)

### 2. Send Email from Outlook
Located in `/sendemailfromoutlook`

A PowerShell script for automated email sending through Outlook.

[View Send Email Documentation](./sendemailfromoutlook/README.md)

## Getting Started

### Prerequisites
- Windows PowerShell 5.1 or later
- Microsoft Outlook (Desktop version)
- Appropriate permissions to run PowerShell scripts in your environment

### Installation
1. Clone this repository:
```powershell
git clone https://github.com/yourusername/outlook-tools.git
```

2. Navigate to the tool directory you want to use:
```powershell
cd outlook-tools/outlooktools
# or
cd outlook-tools/sendemailfromoutlook
```

3. Follow the specific installation instructions in each tool's README.

## General Usage

Each tool has its own specific usage instructions and parameters. Please refer to the individual README files in each tool's directory for detailed documentation.

### Security Note
These scripts interact with Outlook using COM objects. Depending on your organization's security settings, you may need to:
- Unblock the downloaded PS1 files
- Set appropriate PowerShell execution policies
- Handle Outlook security prompts

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Guidelines
1. Maintain consistent formatting with the existing codebase
2. Update documentation for any modified functions
3. Test your changes thoroughly
4. Follow PowerShell best practices

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
```

## Support

If you encounter any issues or have questions:
1. Check the specific tool's README for common issues
2. Open an issue in the GitHub repository
3. Provide relevant details about your environment and the error encountered