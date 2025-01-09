<#
.SYNOPSIS
    Sends personalized emails using Outlook based on a Word template and CSV recipient list.

.DESCRIPTION
    This script automates the process of sending personalized emails through Microsoft Outlook.
    It takes a Word document as a template, replaces placeholders with recipient names,
    and sends individual emails to each recipient listed in a CSV file.

    The script maintains formatting from the original Word document, supports embedded images,
    and provides detailed logging of all operations.

.PARAMETER InputTemplate
    Path to the Word document (.docx) that serves as the email template.
    The template can contain [Name] as a placeholder which will be replaced
    with each recipient's name from the CSV.

.PARAMETER EmailSubject
    The subject line to use for the emails.

.PARAMETER InputCSV
    Path to the CSV file containing recipient information.
    The CSV must have at least two columns: 'Name' and 'Email'.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv"
    Sends personalized emails to all recipients in Input.csv using the template from Sample.docx.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -Verbose
    Sends emails with detailed verbose output showing each step of the process and logs all operations.

.NOTES
    Author: John A. O'Neill Sr.
    Date: 01/08/2025
    Version: 1.1
    Change Date: 01/08/2025
    Change Purpose: Enhanced formatting support and logging

    Prerequisites:
    - Microsoft Office 2016 or later
    - Microsoft Outlook must be installed and configured with a valid email account
    - Microsoft Word must be installed
    - PowerShell 5.1 or later

    Logging:
    - Log files are created in %LOCALAPPDATA%\PowerShell\logs
    - Log files are named SendEmailFromOutlook_YYYY-MM-DD_HH-mm-ss.log
    - Each log entry includes timestamp and operation details

    Required CSV Format:
        Name,GivenName,Surname,Email
        Art Clown,Art,Clown,arttheclown@somedomainsomewhere.com

.LINK
    https://learn.microsoft.com/en-us/office/vba/api/overview/outlook

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None. This script does not generate any output objects.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,
    HelpMessage="Path to the Word document template")]
    [string]$InputTemplate,
    
    [Parameter(Mandatory=$true,
    HelpMessage="Subject line for the emails")]
    [string]$EmailSubject,
    
    [Parameter(Mandatory=$true,
    HelpMessage="Path to the CSV file containing recipient information")]
    [string]$InputCSV
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
    $script:logFilePath = "$log_dir\SendEmailFromOutlook_$time.log"
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

function Initialize-Word {
    [CmdletBinding()]
    param()
    
    try {
        $message = "Initializing Word COM object..."
        Write-Verbose $message
        Write-ToLog $message
        
        # Check if Word is already running
        try {
            $runningWord = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            $message = "Found existing Word instance"
            Write-Verbose $message
            Write-ToLog $message
            return $runningWord
        }
        catch {
            $message = "No existing Word instance found. Creating new instance..."
            Write-Verbose $message
            Write-ToLog $message
        }
        
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $message = "Word COM object initialized successfully"
        Write-Verbose $message
        Write-ToLog $message
        return $word
    }
    catch {
        $errorMessage = "Failed to initialize Word COM object: $_"
        Write-Error $errorMessage
        Write-ToLog "ERROR: $errorMessage"
        throw $errorMessage
    }
}

function Remove-WordInstance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.__ComObject]$WordInstance
    )
    
    try {
        $message = "Cleaning up Word COM object..."
        Write-Verbose $message
        Write-ToLog $message
        
        $WordInstance.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordInstance) | Out-Null
        
        $message = "Word COM object cleaned up successfully"
        Write-Verbose $message
        Write-ToLog $message
    }
    catch {
        $errorMessage = "Failed to clean up Word COM object: $_"
        Write-Error $errorMessage
        Write-ToLog "ERROR: $errorMessage"
    }
}

function Initialize-Outlook {
    [CmdletBinding()]
    param()
    
    try {
        $message = "Initializing Outlook COM object..."
        Write-Verbose $message
        Write-ToLog $message
        
        $outlook = New-Object -ComObject Outlook.Application
        
        # Verify Outlook is properly configured
        try {
            $namespace = $outlook.GetNameSpace("MAPI")
            $namespace.GetDefaultFolder(6) | Out-Null # 6 = olFolderInbox
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        }
        catch {
            throw "Outlook is not properly configured with an email account"
        }
        
        $message = "Outlook COM object initialized successfully"
        Write-Verbose $message
        Write-ToLog $message
        return $outlook
    }
    catch {
        $errorMessage = "Failed to initialize Outlook COM object: $_"
        Write-Error $errorMessage
        Write-ToLog "ERROR: $errorMessage"
        throw $errorMessage
    }
}

function Remove-OutlookInstance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.__ComObject]$OutlookInstance
    )
    
    try {
        $message = "Cleaning up Outlook COM object..."
        Write-Verbose $message
        Write-ToLog $message
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutlookInstance) | Out-Null
        
        $message = "Outlook COM object cleaned up successfully"
        Write-Verbose $message
        Write-ToLog $message
    }
    catch {
        $errorMessage = "Failed to clean up Outlook COM object: $_"
        Write-Error $errorMessage
        Write-ToLog "ERROR: $errorMessage"
    }
}

function Test-FileExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    if (-not (Test-Path $FilePath)) {
        $errorMessage = "File not found: $FilePath"
        Write-Error $errorMessage
        Write-ToLog "ERROR: $errorMessage"
        throw $errorMessage
    }
    
    $message = "File exists: $FilePath"
    Write-Verbose $message
    Write-ToLog $message
}

function Test-CSVFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$CSVPath
    )
    
    try {
        $csv = Import-Csv $CSVPath
        $requiredColumns = @('Name', 'GivenName', 'Surname', 'Email')
        
        foreach ($column in $requiredColumns) {
            if ($csv[0].PSObject.Properties.Name -notcontains $column) {
                throw "CSV file is missing required column: $column"
            }
        }
        
        $message = "CSV format validation passed"
        Write-Verbose $message
        Write-ToLog $message
    }
    catch {
        $errorMessage = "CSV format validation failed: $_"
        Write-Error $errorMessage
        Write-ToLog "ERROR: $errorMessage"
        throw $errorMessage
    }
}

function Get-TemplateContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.__ComObject]$WordInstance,
        
        [Parameter(Mandatory=$true)]
        [string]$TemplatePath
    )
    
    $doc = $null
    
    try {
        Write-ToLog "Starting document processing..."
        
        # Constants for Word
        $wdFormatFilteredHTML = 10
        $wdDoNotSaveChanges = 0
        
        # Create temp file path for HTML
        $tempFile = [System.IO.Path]::GetTempFileName()
        $htmlFile = $tempFile + ".htm"
        Rename-Item -Path $tempFile -NewName ([System.IO.Path]::GetFileName($htmlFile))
        Write-ToLog "Created temporary HTML file: $htmlFile"
        
        # Configure Word instance
        $WordInstance.Visible = $false
        $WordInstance.DisplayAlerts = 0
        
        # Open the document
        Write-ToLog "Opening document..."
        $doc = $WordInstance.Documents.Open($TemplatePath)
        Start-Sleep -Seconds 1
        
        # Configure Web Options
        Write-ToLog "Configuring HTML options..."
        $doc.WebOptions.RelyOnCSS = $true
        $doc.WebOptions.OrganizeInFolder = $false
        $doc.WebOptions.UseLongFileNames = $true
        $doc.WebOptions.RelyOnVML = $false
        $doc.WebOptions.AllowPNG = $true
        $doc.WebOptions.Encoding = 65001 # UTF-8
        
        # Save as filtered HTML
        Write-ToLog "Saving as HTML..."
        $doc.SaveAs([ref]$htmlFile, [ref]$wdFormatFilteredHTML)
        Start-Sleep -Seconds 1
        
        # Close the document before reading the file
        Write-ToLog "Closing document..."
        $doc.Close([ref]$wdDoNotSaveChanges)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        $doc = $null
        
        # Wait for file to be available
        Write-ToLog "Waiting for file to be available..."
        $maxAttempts = 5
        $attempt = 0
        $content = $null
        
        while ($attempt -lt $maxAttempts) {
            try {
                $attempt++
                Write-ToLog "Attempt $attempt to read file..."
                Start-Sleep -Seconds 1
                $content = [System.IO.File]::ReadAllText($htmlFile, [System.Text.Encoding]::UTF8)
                break
            }
            catch {
                if ($attempt -eq $maxAttempts) {
                    throw
                }
                Write-ToLog "File not yet available, retrying..."
                Start-Sleep -Seconds 2
            }
        }
        
        # Clean up
        Write-ToLog "Cleaning up..."
        if (Test-Path $htmlFile) {
            Remove-Item -Path $htmlFile -Force -ErrorAction Stop
        }
        
        Write-ToLog "Document processing completed successfully"
        
        return $content
    }
    catch {
        $errorMessage = "Failed to get template content: $_"
        Write-Error $errorMessage
        Write-ToLog "ERROR: $errorMessage"
        
        # Cleanup in case of error
        if ($doc) {
            try {
                $doc.Close([ref]$wdDoNotSaveChanges)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
            catch {
                Write-ToLog "Warning: Error during document cleanup: $_"
            }
        }
        
        if (Test-Path $htmlFile) {
            try {
                Remove-Item -Path $htmlFile -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-ToLog "Warning: Could not remove temporary file: $_"
            }
        }
        
        throw $errorMessage
    }
}

function Send-PersonalizedEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.__ComObject]$OutlookInstance,
        
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [string]$To,
        
        [Parameter(Mandatory=$true)]
        [string]$Body,
        
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$RecipientData
    )
    
    $mail = $null
    
    try {
        $message = "Creating email for recipient: $To"
        Write-Verbose $message
        Write-ToLog $message
        
        $mail = $OutlookInstance.CreateItem(0)
        Write-ToLog "Created mail item"
        
        # Basic properties
        $mail.Subject = $Subject
        $mail.To = $To
        
        # Ensure proper HTML formatting
        Write-ToLog "Setting email format and content..."
        $mail.BodyFormat = 2  # olFormatHTML
        
        # Personalize content with recipient data
        $personalizedBody = $Body
        $personalizedBody = $personalizedBody -replace '\[GivenName\]', $RecipientData.GivenName
        
        # Clean up the HTML if needed
        $cleanBody = $personalizedBody
        if (-not $personalizedBody.Contains("<!DOCTYPE")) {
            $cleanBody = @"
<!DOCTYPE html>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<style>
body { font-family: Calibri, Arial, sans-serif; }
</style>
</head>
<body>
$personalizedBody
</body>
</html>
"@
        }
        
        $mail.HTMLBody = $cleanBody
        
        Write-ToLog "Sending email..."
        $mail.Send()
        
        $message = "Email sent successfully to $To"
        Write-Host $message -ForegroundColor Green
        Write-ToLog $message
    }
    catch {
        $errorMessage = "Failed to send email to ${To}: $_"
        Write-Warning $errorMessage
        Write-ToLog "WARNING: $errorMessage"
    }
    finally {
        if ($mail) {
            try {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null
                Write-ToLog "Released Mail COM object"
            }
            catch {
                Write-ToLog "Warning: Error releasing mail object: $_"
            }
        }
    }
}

# Main script execution
try {
    # Create new log file for this run
    New-Log
    Write-ToLog "Script started with parameters: InputTemplate=$InputTemplate, EmailSubject=$EmailSubject, InputCSV=$InputCSV"
    
    # Validate input files
    Test-FileExists -FilePath $InputTemplate
    Test-FileExists -FilePath $InputCSV
    
    # Validate CSV format
    Test-CSVFormat -CSVPath $InputCSV
    
    # Initialize COM objects
    $word = Initialize-Word
    $outlook = Initialize-Outlook
    
    # Get template content
    $templateContent = Get-TemplateContent -WordInstance $word -TemplatePath $InputTemplate
    
    # Import CSV
    $message = "Importing CSV file: $InputCSV"
    Write-Verbose $message
    Write-ToLog $message
    $recipients = Import-Csv $InputCSV
    Write-ToLog "Found $($recipients.Count) recipients in CSV file"
    
    # Process each recipient
    $totalEmails = $recipients.Count
    $currentEmail = 0
    
    foreach ($recipient in $recipients) {
        $currentEmail++
        $progressMessage = "Processing email $currentEmail of $totalEmails"
        Write-Progress -Activity "Sending Emails" -Status $progressMessage -PercentComplete (($currentEmail / $totalEmails) * 100)
        Write-ToLog $progressMessage
        
        # Send email with recipient data
        Send-PersonalizedEmail -OutlookInstance $outlook -Subject $EmailSubject -To $recipient.Email -Body $templateContent -RecipientData $recipient
        
        # Small delay between emails to prevent throttling
        Start-Sleep -Milliseconds 500
    }
    
    $message = "Email sending process completed!"
    Write-Host "`n$message" -ForegroundColor Green
    Write-ToLog $message
}
catch {
    $errorMessage = "A critical error occurred: $_"
    Write-Error $errorMessage
    Write-ToLog "ERROR: $errorMessage"
    exit 1
}
finally {
    # Clean up COM objects
    if ($word) { 
        Remove-WordInstance -WordInstance $word 
    }
    if ($outlook) { 
        Remove-OutlookInstance -OutlookInstance $outlook 
    }
    
    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-ToLog "Script execution completed"
}