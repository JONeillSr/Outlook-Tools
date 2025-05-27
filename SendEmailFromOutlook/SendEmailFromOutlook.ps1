<#
.SYNOPSIS
    Sends personalized emails using Outlook based on a Word template and CSV recipient list.

.DESCRIPTION
    This script automates the process of sending personalized emails through Microsoft Outlook.
    It takes a Word document as a template, replaces placeholders with recipient names,
    and sends individual emails to each recipient listed in a CSV file.

    The script maintains formatting from the original Word document, supports embedded images,
    allows attachments, provides detailed logging of all operations, and includes options
    for setting email importance, delivery receipts, and read receipts.

.PARAMETER InputTemplate
    Path to the Word document (.docx) that serves as the email template.
    The template can contain [Name] as a placeholder which will be replaced
    with each recipient's name from the CSV.

.PARAMETER EmailSubject
    The subject line to use for the emails.

.PARAMETER InputCSV
    Path to the CSV file containing recipient information.
    The CSV must have at least two columns: 'Name' and 'Email'.

.PARAMETER FromAddress
    Optional. The email address to use as the sender. If not specified, 
    uses the default Outlook account. The specified address must be configured 
    in Outlook as a valid sending account.

.PARAMETER AttachmentPath
    Optional. Path to a file that will be attached to each email sent.
    If specified, the file must exist and will be attached to every email.

.PARAMETER HighImportance
    Optional. Sets the email importance to high priority.

.PARAMETER DeliveryReceipt
    Optional. Requests delivery receipts for all emails sent.

.PARAMETER ReadReceipt
    Optional. Requests read receipts for all emails sent.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv"
    Sends personalized emails to all recipients in Input.csv using the template from Sample.docx with the default sender address.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -FromAddress "example@somedomainsomewhere.com"
    Sends emails from the specified address instead of the default Outlook account.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -AttachmentPath "C:\Temp\Brochure.pdf"
    Sends emails with the specified PDF file attached to each email.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -FromAddress "example@somedomainsomewhere.com" -AttachmentPath "C:\Temp\Brochure.pdf"
    Sends emails from a custom address with an attachment included.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -Verbose
    Sends emails with detailed verbose output showing each step of the process and logs all operations.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -FromAddress "example@somedomainsomewhere.com"
    Sends an email with a custom From address.
    
.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -FromAddress "example@somedomainsomewhere.com" -AttachmentPath "C:\Temp\Brochure.pdf"
    Sends an email with a custom From address and includes an attachment.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Hello in 2025!" -InputCSV "C:\Temp\Input.csv" -AttachmentPath "C:\Temp\Brochure.pdf"
    Sends an email with the default From address and includes an attachment.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Quarterly Report" -InputCSV "C:\Temp\Input.csv" -HighImportance -DeliveryReceipt
    Sends high importance emails with delivery receipt requests.

.EXAMPLE
    .\SendEmailFromOutlook.ps1 -InputTemplate "C:\Temp\Sample.docx" -EmailSubject "Training Materials" -InputCSV "C:\Temp\Input.csv" -ReadReceipt -AttachmentPath "C:\Temp\Training.pdf"
    Sends emails with read receipt requests and includes a training PDF attachment.

.NOTES
    Author: John A. O'Neill Sr.
    Date: 01/08/2025
    Version: 1.4
    Change Date: 05/27/2025
    Change Purpose: Added ability to set importance, delivery, and read receipt options.

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

    From Address Notes:
    - The FromAddress parameter must be an email address configured in your Outlook
    - If the address is not configured in Outlook, the email will fail to send
    - If FromAddress is not specified, the default Outlook account will be used
    - If the FromAddress is an alias, and the alias doesn't exist or Exhange isn't configured to allow send from alias, the primary account will be used to send the email

    Attachment Notes:
    - The AttachmentPath parameter must point to an existing file
    - The same attachment will be included in every email sent
    - Large attachments may cause slower sending or delivery issues

.LINK
    https://learn.microsoft.com/en-us/office/vba/api/overview/outlook

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None. This script does not generate any output objects.
#>

param(
    [Parameter(Mandatory=$true,
    HelpMessage="Path to the Word document template")]
    [string]$InputTemplate,
    
    [Parameter(Mandatory=$true,
    HelpMessage="Subject line for the emails")]
    [string]$EmailSubject,
    
    [Parameter(Mandatory=$true,
    HelpMessage="Path to the CSV file containing recipient information")]
    [string]$InputCSV,
    
    [Parameter(Mandatory=$false,
    HelpMessage="Email address to use as sender (must be configured in Outlook)")]
    [string]$FromAddress = $null,
    
    [Parameter(Mandatory=$false,
    HelpMessage="Path to file that will be attached to each email")]
    [string]$AttachmentPath = $null,

    [Parameter(Mandatory=$false,
    HelpMessage="Set email importance to high")]
    [switch]$HighImportance,
    
    [Parameter(Mandatory=$false,
    HelpMessage="Request delivery receipt for emails")]
    [switch]$DeliveryReceipt,
    
    [Parameter(Mandatory=$false,
    HelpMessage="Request read receipt for emails")]
    [switch]$ReadReceipt
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

function Test-AttachmentPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$AttachmentPath
    )
    
    if ([string]::IsNullOrEmpty($AttachmentPath)) {
        $message = "No attachment specified"
        Write-Verbose $message
        Write-ToLog $message
        return $false
    }
    
    try {
        if (-not (Test-Path $AttachmentPath)) {
            $errorMessage = "Attachment file not found: $AttachmentPath"
            Write-Error $errorMessage
            Write-ToLog "ERROR: $errorMessage"
            throw $errorMessage
        }
        
        # Check if it's a file (not a directory)
        $item = Get-Item $AttachmentPath
        if ($item.PSIsContainer) {
            $errorMessage = "Attachment path is a directory, not a file: $AttachmentPath"
            Write-Error $errorMessage
            Write-ToLog "ERROR: $errorMessage"
            throw $errorMessage
        }
        
        # Get file size and warn if it's large
        $fileSizeMB = [math]::Round($item.Length / 1MB, 2)
        if ($fileSizeMB -gt 10) {
            $warningMessage = "Attachment file is large ($fileSizeMB MB). This may cause delivery issues."
            Write-Warning $warningMessage
            Write-ToLog "WARNING: $warningMessage"
        }
        
        $message = "Attachment validation passed: $AttachmentPath ($fileSizeMB MB)"
        Write-Verbose $message
        Write-ToLog $message
        return $true
    }
    catch {
        $errorMessage = "Attachment validation failed: $_"
        Write-Error $errorMessage
        Write-ToLog "ERROR: $errorMessage"
        throw $errorMessage
    }
}

function Test-FromAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.__ComObject]$OutlookInstance,
        
        [Parameter(Mandatory=$false)]
        [string]$FromAddress
    )
    
    if ([string]::IsNullOrEmpty($FromAddress)) {
        $message = "No custom FromAddress specified, using default Outlook account"
        Write-Verbose $message
        Write-ToLog $message
        return @{
            IsValid = $true
            Account = $null
            Address = $null
            IsAlias = $false
        }
    }
    
    try {
        $message = "Validating FromAddress: $FromAddress"
        Write-Verbose $message
        Write-ToLog $message
        
        $namespace = $OutlookInstance.GetNameSpace("MAPI")
        $accounts = $namespace.Accounts
        
        # First check if it's a direct account match
        $foundAccount = $null
        for ($i = 1; $i -le $accounts.Count; $i++) {
            $account = $accounts.Item($i)
            if ($account.SmtpAddress -eq $FromAddress) {
                $foundAccount = $account
                break
            }
        }
        
        if ($foundAccount) {
            $message = "FromAddress validation passed (direct account): $FromAddress"
            Write-Verbose $message
            Write-ToLog $message
            return @{
                IsValid = $true
                Account = $foundAccount
                Address = $FromAddress
                IsAlias = $false
            }
        }
        
        # If not found as direct account, assume it's an alias
        # For aliases, we'll use the primary account but set the SentOnBehalfOfName
        $primaryAccount = $accounts.Item(1)  # Use first/primary account
        
        $message = "FromAddress '$FromAddress' appears to be an alias. Will use primary account '$($primaryAccount.SmtpAddress)' with SentOnBehalfOfName"
        Write-Verbose $message
        Write-ToLog $message
        
        return @{
            IsValid = $true
            Account = $primaryAccount
            Address = $FromAddress
            IsAlias = $true
        }
    }
    catch {
        $errorMessage = "Failed to validate FromAddress: $_"
        Write-Warning $errorMessage
        Write-ToLog "WARNING: $errorMessage"
        return @{
            IsValid = $false
            Account = $null
            Address = $null
            IsAlias = $false
        }
    }
    finally {
        if ($namespace) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        }
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
        [PSCustomObject]$RecipientData,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$FromAddressInfo = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$AttachmentPath = $null,

        [Parameter(Mandatory=$false)]
        [bool]$SetHighImportance = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$RequestDeliveryReceipt = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$RequestReadReceipt = $false
    )
    
    $mail = $null
    
    try {
        $message = "Creating email for recipient: $To"
        Write-Verbose $message
        Write-ToLog $message
        
        $mail = $OutlookInstance.CreateItem(0)
        Write-ToLog "Created mail item"
        
        # Handle From address configuration
        if ($FromAddressInfo -and $FromAddressInfo.IsValid) {
            if ($FromAddressInfo.IsAlias) {
                # For aliases, use SentOnBehalfOfName and set the From property
                if ($FromAddressInfo.Account) {
                    $mail.SendUsingAccount = $FromAddressInfo.Account
                }
                
                # Set the From property to the alias address
                $mail.SentOnBehalfOfName = $FromAddressInfo.Address
                
                # Alternative approach: directly set From property (works in some Exchange configurations)
                try {
                    $mail.From = $FromAddressInfo.Address
                }
                catch {
                    Write-ToLog "Note: Could not set From property directly: $_"
                }
                
                $message = "Set From address to alias: $($FromAddressInfo.Address)"
                Write-Verbose $message
                Write-ToLog $message
            }
            else {
                # For direct accounts, use SendUsingAccount
                $mail.SendUsingAccount = $FromAddressInfo.Account
                $message = "Set From address to account: $($FromAddressInfo.Address)"
                Write-Verbose $message
                Write-ToLog $message
            }
        }
        
        # Basic properties
        $mail.Subject = $Subject
        $mail.To = $To
        
        if ($SetHighImportance) {
            $mail.Importance = 2  # olImportanceHigh (0=Low, 1=Normal, 2=High)
            Write-ToLog "Set email importance to High"
        }
        
        if ($RequestDeliveryReceipt) {
            try {
                # For delivery receipts, we need to set the OriginatorDeliveryReportRequested property
                # This is done through MAPI properties
                $mail.OriginatorDeliveryReportRequested = $true
                Write-ToLog "Requested delivery receipt"
            }
            catch {
                try {
                    # Alternative method using PropertyAccessor for delivery receipt
                    $propertyAccessor = $mail.PropertyAccessor
                    # MAPI property for delivery receipt request
                    $deliveryReceiptProperty = "http://schemas.microsoft.com/mapi/proptag/0x23000003"
                    $propertyAccessor.SetProperty($deliveryReceiptProperty, $true)
                    Write-ToLog "Requested delivery receipt (via PropertyAccessor)"
                }
                catch {
                    Write-ToLog "WARNING: Could not set delivery receipt request: $_"
                    Write-Warning "Delivery receipt request not supported in this Outlook configuration"
                }
            }
        }
        
        if ($RequestReadReceipt) {
            try {
                $mail.ReadReceiptRequested = $true
                Write-ToLog "Requested read receipt"
            }
            catch {
                Write-ToLog "WARNING: Could not set read receipt request: $_"
                Write-Warning "Read receipt request not supported in this Outlook configuration"
            }
        }
        
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
        
        # Add attachment if specified
        if (-not [string]::IsNullOrEmpty($AttachmentPath)) {
            try {
                Write-ToLog "Adding attachment: $AttachmentPath"
                $attachment = $mail.Attachments.Add($AttachmentPath)
                $attachmentName = [System.IO.Path]::GetFileName($AttachmentPath)
                $message = "Attachment added successfully: $attachmentName"
                Write-Verbose $message
                Write-ToLog $message
            }
            catch {
                $errorMessage = "Failed to add attachment '$AttachmentPath': $_"
                Write-Warning $errorMessage
                Write-ToLog "WARNING: $errorMessage"
                # Don't throw - continue sending email without attachment
            }
        }
        
        Write-ToLog "Sending email..."
        $mail.Send()
        
        $fromInfo = if ($FromAddressInfo -and $FromAddressInfo.IsValid) { 
            $aliasInfo = if ($FromAddressInfo.IsAlias) { " (alias)" } else { "" }
            " from $($FromAddressInfo.Address)$aliasInfo" 
        } else { 
            "" 
        }
        $attachmentInfo = if (-not [string]::IsNullOrEmpty($AttachmentPath)) { 
            " with attachment" 
        } else { 
            "" 
        }

        # Build options info string
        $optionsInfo = @()
        if ($SetHighImportance) { $optionsInfo += "high importance" }
        if ($RequestDeliveryReceipt) { $optionsInfo += "delivery receipt" }
        if ($RequestReadReceipt) { $optionsInfo += "read receipt" }
        $optionsString = if ($optionsInfo.Count -gt 0) { " (" + ($optionsInfo -join ", ") + ")" } else { "" }
        
        $message = "Email sent successfully to $To$fromInfo$attachmentInfo$optionsString"
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
    # Create a new log file
    New-Log
    $fromAddressLog = if ($FromAddress) { ", FromAddress=$FromAddress" } else { "" }
    $attachmentLog = if ($AttachmentPath) { ", AttachmentPath=$AttachmentPath" } else { "" }
    $optionsLog = @()
    if ($HighImportance) { $optionsLog += "HighImportance" }
    if ($DeliveryReceipt) { $optionsLog += "DeliveryReceipt" }
    if ($ReadReceipt) { $optionsLog += "ReadReceipt" }
    $optionsString = if ($optionsLog.Count -gt 0) { ", Options=" + ($optionsLog -join ",") } else { "" }

    Write-ToLog "Script started with parameters: InputTemplate=$InputTemplate, EmailSubject=$EmailSubject, InputCSV=$InputCSV$fromAddressLog$attachmentLog$optionsString"

    # Validate input files
    Test-FileExists -FilePath $InputTemplate
    Test-FileExists -FilePath $InputCSV
    
    # Validate CSV format
    Test-CSVFormat -CSVPath $InputCSV
    
    # Validate attachment if specified
    $hasAttachment = Test-AttachmentPath -AttachmentPath $AttachmentPath
    
    # Initialize COM objects
    $word = Initialize-Word
    $outlook = Initialize-Outlook
    
    # Validate and get FromAddress account if specified
    $fromAddressInfo = Test-FromAddress -OutlookInstance $outlook -FromAddress $FromAddress
    
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
        
        # Send email with recipient data and all options
        Send-PersonalizedEmail -OutlookInstance $outlook -Subject $EmailSubject -To $recipient.Email -Body $templateContent -RecipientData $recipient -FromAddressInfo $fromAddressInfo -AttachmentPath $AttachmentPath -SetHighImportance $HighImportance -RequestDeliveryReceipt $DeliveryReceipt -RequestReadReceipt $ReadReceipt
        
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