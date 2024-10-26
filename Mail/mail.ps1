# Install Import-Excel Modul in PS

# Change values
param (
    [string]$ExcelFile, # needs to be specified
    [string]$FileDirectory, # needs to be specified
    [string]$LogFile, # if empty logfile gets created in this path
    [string]$UserName, # if emty its the email
    [string]$SenderEmail, # needs to be specified
    [string]$SmtpServer, # needs to be specified
    [int]$SmtpPort, # needs to be specified
    [string]$MailSubject = "Mail automation",
    [string]$MailText = "Hello [name]!"
)

if (-not $LogFile) { # If log file is not specified, create it in the current directory
    $LogFile = Join-Path -Path (Get-Location) -ChildPath "log.txt"
}

if (-not $userName) { # If user name is not specified, use the sender email
    $userName = $senderEmail
}

# example call:
# .\mailParams.ps1 -ExcelFile "C:\Users\julia\Desktop\Names.xlsx" -FileDirectory "C:\Users\julia\Desktop\Files" -SenderEmail "jh4112000@gmx.de" -SmtpServer "mail.gmx.net" -SmtpPort 587

# Log file
'' | Out-File -FilePath $LogFile -Force

$successCount = 0
$failureCount = 0

# Script start ########################################

Write-Host "Script for mail automation started`n"
Write-Host "Authentification loads`n"

# Load Excel file
$excel = Import-Excel $excelFile
$rowCount = 0

# SMTP-settings
$smtpCredential = Get-Credential -Message "Enter your credentials" -UserName "$userName"

# Iterate over each row in the Excel file
foreach ($row in $excel) {
    # Get name and email from the current row
    $name = $row.Name
    $email = $row.Email

    # Skip iteration if one of the variables is empty
    if (-not $name -or -not $email) {
        continue
    }
    $rowCount++ # Increase row count
    
    # Construct the file path
    $filePath = Join-Path $fileDirectory "$name.pdf"

    # Check if the file exists
    if (-not (Test-Path $filePath)) {
        $failureCount++
        "[FAILED] Email to $email failed: File not found for $name on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -Append -FilePath $logFile
        continue # Skip to the next iteration
    }

    # Check if the SMTP server is reachable
    if (-not (Test-Connection -ComputerName $smtpServer -Count 1 -Quiet)) {
        $failureCount++
        "[FAILED] Email to $email failed: Email server not reachable for $name on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -Append -FilePath $logFile
        continue # Skip to the next iteration
    }

    # Construct the email message
    $subject = $mailSubject
    $attachment = $filePath
    $message = @{
        From       = $senderEmail
        To         = $email
        Subject    = $subject
        Body       = $mailText.Replace("[name]", $name)
        Attachments = $attachment
        SmtpServer = $smtpServer
        Port       = $smtpPort
        UseSSL     = $true
        Credential = $smtpCredential
    }

    # Try to send the email
    try { 
        Send-MailMessage @message -ErrorAction Stop
        $successCount++
        # Log success
        "[SUCCESS] Email to $email sent successfully for $name on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -Append -FilePath $logFile
    } catch { 
        $failureCount++
        "[FAILED] Email to $email failed: Authentication error for $name on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'). Error: $_" | Out-File -Append -FilePath $logFile
    }
}

# Give feedback to the user
Write-Host "`RESULT:`nSuccess: $successCount from $rowCount - Failed: $failureCount from $rowCount.`nSee log.txt for more information."
Write-Host "`nScript has been completed."

