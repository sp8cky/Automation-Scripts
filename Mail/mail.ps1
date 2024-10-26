# Change values
param (
    [string]$ExcelFile, # Required
    [string]$FileDirectory, # Required
    [string]$SenderEmail, # Required
    [string]$SmtpServer, # Required
    [int]$SmtpPort, # Required
    [string]$LogFile, # if empty logfile gets created in this path
    [string]$UserName, # if empty its the email
    [string]$MailSubject = "Your file",
    [string]$MailText = "Hello [name]!"
)

# check if ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

if (-not $UserName) { # If user name is not specified, use the sender email
    $UserName = $SenderEmail
}

# Initialize log file
try {
    if (-not $LogFile) {
        # If no logfile path is specified, create log.txt in the current directory
        $LogFile = Join-Path -Path (Get-Location) -ChildPath "log.txt"
    } elseif (-not $LogFile.EndsWith(".txt")) {
        # If just a directory is specified, create log.txt in that directory
        $LogFile = Join-Path -Path $LogFile -ChildPath "log.txt"
    }
    '' | Out-File -FilePath $LogFile -Force # Initialize log file
} catch {
    Write-Host "Error initializing log file: $_"
    exit 1
}

$successCount = 0
$failureCount = 0

# Script start ########################################

Write-Host "Script for mail automation started`n"
Write-Host "Authentication loads`n"

# Load Excel file
$excel = Import-Excel $ExcelFile
$rowCount = 0

# SMTP settings
$smtpCredential = Get-Credential -Message "Enter your credentials" -UserName "$UserName"

# Iterate over each row in the Excel file
foreach ($row in $excel) {
    # Get name and email from the current row
    $name = $row.Name
    $email = $row.Email

    # Skip iteration if one of the variables is empty
    if (-not $name -or -not $email) {
        continue
    }
    $rowCount++
    
    # Construct the file path pattern for the file (any extension)
    $filePathPattern = Join-Path -Path $FileDirectory -ChildPath "$name.*"
    $filePath = Get-ChildItem -Path $filePathPattern -File -ErrorAction SilentlyContinue | Select-Object -First 1

    # Check if the file exists
    if (-not $filePath) {
        $failureCount++
        "[FAILED] Email to $email failed: File not found for $name on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -Append -FilePath $LogFile
        continue
    }

    # Check if the SMTP server is reachable
    if (-not (Test-Connection -ComputerName $SmtpServer -Count 1 -Quiet)) {
        $failureCount++
        "[FAILED] Email to $email failed: Email server not reachable for $name on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -Append -FilePath $LogFile
        continue
    }

    # Construct the email message
    $subject = $MailSubject
    $attachment = $filePath.FullName
    $message = @{
        From       = $SenderEmail
        To         = $email
        Subject    = $subject
        Body       = $MailText.Replace("[name]", $name)
        Attachments = $attachment
        SmtpServer = $SmtpServer
        Port       = $SmtpPort
        UseSSL     = $true
        Credential = $smtpCredential
    }

    # Try to send the email
    try { 
        Send-MailMessage @message -ErrorAction Stop
        $successCount++
        # Log success
        "[SUCCESS] Email to $email sent successfully for $name on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -Append -FilePath $LogFile
    } catch { 
        $failureCount++
        "[FAILED] Email to $email failed: Authentication error for $name on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'). Error: $_" | Out-File -Append -FilePath $LogFile
    }
}

# Give feedback to the user
Write-Host "`RESULT:`nSuccess: $successCount from $rowCount - Failed: $failureCount from $rowCount.`nSee log.txt for more information."
Write-Host "`nScript has been completed."