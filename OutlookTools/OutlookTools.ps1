<#
.SYNOPSIS
    PowerShell script to enumerate folders in an Outlook mailbox or retrieve email addresses from a specified folder.

.DESCRIPTION
    This script provides these functionalities:
    1. Enumerate all folders in an Outlook mailbox, including item counts.
    2. Retrieve sender names and email addresses from a specific folder in the mailbox.

.PARAMETER GetMailboxFolders
    Call this parameter to enumerate all folders in the mailbox.

.PARAMETER GetEmailsFromFolder
    Call this parameter to retrieve email addresses from a specified folder.

.PARAMETER FolderPath
    Specify the folder path for the GetEmailsFromFolder function. (Required for email retrieval).

.PARAMETER MailboxName
    Specify the mailbox display name (e.g., "Mailbox - Your Name"). Defaults to the primary mailbox.

.EXAMPLE
    .\OutlookScript.ps1 -GetMailboxFolders -OutputFolder "C:\Reports"


.EXAMPLE
    .\OutlookScript.ps1 -GetMailboxFolders -MailboxName "Mailbox - Art Clown" -OutputFolder "C:\Reports"

    Retrieves folders recursively for all mailboxes configured with Outlook and the results are combined and then saved in the "C:\Reports" folder.

.EXAMPLE
    .\OutlookScript.ps1 -GetEmailsFromFolder -FolderPath "Inbox\MyFolder" -MailboxName "Mailbox - Art Clown" -OutputFolder "D:\Exports"

    Retrieves email addresses from the folder "Inbox\MyFolder" in the mailbox "Mailbox - Art Clown", then saves the output in the "D:\Exports" folder.

.EXAMPLE
    .\OutlookScript.ps1 -GetEmailsFromFolder -FolderPath "Inbox\MyFolder" -MailboxName "Mailbox - Art Clown"

    Retrieves email addresses from the folder "Inbox\MyFolder" in the mailbox "Mailbox - Art Clown", then saves the output on the user's Desktop.

.NOTES
    Author: John A. O'Neill Sr.
    Date: 01/08/2025
    Version: 1.0
    Change Date:
    Change Purpose:

    Prerequisites:
    - Microsoft Outlook must be installed and configured with a valid email account
    - PowerShell 5.1 or later

    Logging:
    - Log files are created in %LOCALAPPDATA%\PowerShell\logs
    - Log files are named OutlookTools_YYYY-MM-DD_HH-mm-ss.log
    - Each log entry includes timestamp and operation details

.LINK
    https://learn.microsoft.com/en-us/office/vba/api/overview/outlook

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None. This script does not generate any output objects.
#>

[CmdletBinding()]
param (
    [switch]$GetMailboxFolders,
    [switch]$GetEmailsFromFolder,
    [string]$FolderPath,
    [string]$MailboxName,
    [string]$OutputFolder = "$env:USERPROFILE\Desktop" # Default to Desktop
)

# Initial variables defined at script scope
$script:scriptpath = $env:localappdata
$script:dir = "$scriptpath\PowerShell"
$script:log_dir = "$dir\logs"
$script:logFilePath = $null

# Ensure log directory exists
if (-not (Test-Path -Path $log_dir)) {
    New-Item -Path $log_dir -ItemType Directory
}

# Create a new log file with a timestamp in the filename
Function New-Log {
    $time = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $script:logFilePath = "$log_dir\OutlookTools_$time.log"
    New-Item -ItemType File $script:logFilePath -Force -ErrorAction SilentlyContinue > $null
    Write-Host "Created new log file $script:LogFilePath" -ForegroundColor DarkYellow
}

# Write data to log file
Function Write-ToLog {
    param (
        [string]$message
    )
    if (-not $script:logFilePath) {
        Write-Warning "Log file path not initialized. Creating new log file."
        New-Log
    }
    $time = Get-Date
    $logEntryStr = "$time`t" + $message
    $logEntryStr | Out-File -Append $script:logFilePath -Width 400
}

# Function to initialize Outlook COM object
function Initialize-Outlook {
    Write-Verbose "Initializing Outlook COM object..."
    Write-ToLog -message "Initializing Outlook COM object..."
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    return $namespace
}

# Function to resolve the output file name (handle overwrites or create a new file)
function Resolve-OutputFile {
    param (
        [string]$FilePath
    )

    if (Test-Path -Path $FilePath) {
        Write-Host "The file '$FilePath' already exists." -ForegroundColor Yellow
        Write-ToLog -message "The file '$FilePath' already exists."
        $response = Read-Host "Do you want to overwrite it? (Y/N)"
        if ($response -like "N*") {
            # Create a new unique file name
            $timestamp = Get-Date -Format "yyyyMMddHHmmss"
            $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath) + "_$timestamp" + [System.IO.Path]::GetExtension($FilePath)
            $FilePath = Join-Path -Path ([System.IO.Path]::GetDirectoryName($FilePath)) -ChildPath $newFileName
            Write-Host "A new file will be created: '$FilePath'" -ForegroundColor Green
            Write-ToLog -message "A new file will be created: '$FilePath'"
        } else {
            Write-Host "The file will be overwritten." -ForegroundColor Green
            Write-ToLog -message "The file will be overwritten."
        }
    }
    return $FilePath
}

# Function to get all folders in one or all mailboxes
function Get-MailboxFolders {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Namespace,
        [string]$MailboxName = ""
    )

    # Recursive folder retrieval
    function Get-FolderList {
        param (
            [object]$Folder,
            [string]$ParentPath = "",
            [string]$Mailbox
        )

        # Construct FolderPath relative to the mailbox root
        $relativePath = if ($ParentPath) {
            "$ParentPath\$($Folder.Name)"
        } else {
            "\$($Folder.Name)"
        }

        # Strip the mailbox name from the FolderPath (if present)
        if ($relativePath.StartsWith("\$Mailbox")) {
            $relativePath = $relativePath.Substring($Mailbox.Length + 1)
        }

        # Output mailbox and folder path as separate columns
        [PSCustomObject]@{
            Mailbox    = $Mailbox
            FolderPath = $relativePath
            ItemCount  = $Folder.Items.Count
        }

        # Process subfolders recursively
        foreach ($subfolder in $Folder.Folders) {
            Get-FolderList -Folder $subfolder -ParentPath $relativePath -Mailbox $Mailbox
        }
    }

    # If a mailbox name is specified, get its root folder
    if ($MailboxName) {
        $rootFolder = $Namespace.Folders.Item($MailboxName)
        if (-not $rootFolder) {
            Write-Error "Mailbox '$MailboxName' not found."
            Write-ToLog -message "Mailbox '$MailboxName' not found."
            return
        }
        Get-FolderList -Folder $rootFolder -Mailbox $MailboxName
    } else {
        # Iterate through all mailboxes
        $mailboxes = $Namespace.Folders
        foreach ($mailbox in $mailboxes) {
            Get-FolderList -Folder $mailbox -Mailbox $mailbox.Name
        }
    }
}

# Function to retrieve email addresses from a folder
function Get-EmailsFromFolder {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Namespace,
        [string]$FolderPath,
        [string]$MailboxName
    )

    # Get the root folder of the specified mailbox
    $rootFolder = $Namespace.Folders.Item($MailboxName)
    if (-not $rootFolder) {
        Write-Error "Mailbox '$MailboxName' not found."
        Write-ToLog -message "The specified output folder '$OutputFolder' does not exist."
        return
    }

    # Navigate to the target folder
    $folder = $rootFolder
    $folderNames = $FolderPath.TrimStart("\").Split("\")
    if ($folderNames[0] -eq $MailboxName) {
        Write-Verbose "Skipping mailbox name: '$MailboxName'"
        Write-ToLog -message "Skipping mailbox name: '$MailboxName'"
        $folderNames = $folderNames[1..$folderNames.Length] # Skip the mailbox name
    }

    foreach ($folderName in $folderNames) {
        Write-Verbose "Looking for folder: '$folderName' in '$($folder.Name)'"
        Write-ToLog -message "Looking for folder: '$folderName' in '$($folder.Name)'"
        try {
            $folder = $folder.Folders.Item($folderName)
            if (-not $folder) {
                Write-Error "Folder '$folderName' not found in path '$FolderPath'."
                Write-ToLog -message "Folder '$folderName' not found in path '$FolderPath'."
                return
            }
        } catch {
            Write-Error "Folder '$folderName' not found or inaccessible in path '$FolderPath'."
            Write-ToLog -message "Folder '$folderName' not found or inaccessible in path '$FolderPath'."
            return
        }
    }

    # Extract email addresses
    $emails = @()
    foreach ($item in $folder.Items) {
        if ($item.Class -eq 43) { # 43 = MailItem
            $emailAddress = $item.SenderEmailAddress
            $displayName = $item.SenderName
            if ($emailAddress -and -not ($emails | Where-Object { $_.Email -eq $emailAddress })) {
                $emails += [PSCustomObject]@{
                    Name  = $displayName
                    Email = $emailAddress
                }
                Write-ToLog -message "Exporting data for $emailAddress"
            }
        }
    }

    # Return the extracted emails
    return $emails
}

# Main Script Logic
New-Log

# Ensure the output folder exists
if (-not (Test-Path -Path $OutputFolder)) {
    Write-Error "The specified output folder '$OutputFolder' does not exist."
    Write-ToLog -message "The specified output folder '$OutputFolder' does not exist."
    return
}

# Initialize Outlook namespace
$namespace = Initialize-Outlook

if ($GetMailboxFolders) {
    $outputFile = Join-Path -Path $OutputFolder -ChildPath "OutlookFolderList.csv"
    $outputFile = Resolve-OutputFile -FilePath $outputFile
    $folders = Get-MailboxFolders -Namespace $namespace -MailboxName $MailboxName
    $folders | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "Folder list saved to $outputFile" -ForegroundColor Green
    Write-ToLog -message "Folder list saved to $outputFile"
}

if ($GetEmailsFromFolder) {
    if (-not $FolderPath) {
        Write-Error "The FolderPath parameter is required for GetEmailsFromFolder."
        Write-ToLog -message "The FolderPath parameter is required for GetEmailsFromFolder."
        return
    }

    $outputFile = Join-Path -Path $OutputFolder -ChildPath "ExtractedEmails.csv"
    $outputFile = Resolve-OutputFile -FilePath $outputFile
    $emails = Get-EmailsFromFolder -Namespace $namespace -FolderPath $FolderPath -MailboxName $MailboxName
    $emails | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "Email addresses saved to $outputFile" -ForegroundColor Green
    Write-ToLog -message "Email addresses saved to $outputFile"
}
