# Mail automation script
This PowerShell script automates the sending of emails with attachments based on a list of recipients stored in an Excel file. It checks whether the files to be sent exist and logs the success or failure of each email transmission in a log file.

## Table of contents
1. [Implemented Features](#Implemented-Features)
2. [Installation and usage](#installation-and-usage)
3. [Disclaimer](#disclaimer)

## Implemented features
- Reads the recipient information (name and e-mail address) from a specified Excel file.
- Sends e-mails with attached files from a specified directory.
- Logs the sending status in a log file.
- Supports user-defined SMTP server settings.

### Project status
- Finished

## Installation and usage
### Requirements
- **PowerShell**: This script requires PowerShell 5.1 or higher.

#### Excel file
The excel file need two columns (every entry one row):
- Name: The name of the recipient.
- Email: The recipient's email address.

| Name  | Email  |
| ---   |  --- |
| john  | john.doe@test.com |
| max  | max@gmail.com |

Clone the repository and navigate into the directory
```
git clone https://github.com/sp8cky/Automation-Scripts/Mail && cd Automation-Scripts/Mail
```
Install dependencies
```
pip install -r requirements.txt
```

### Call 
Change the execution rights (see [here](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4) for more information)
```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
``` 
``` 
.\mail.ps1 -ExcelFile "<Path to excel file>" -FileDirectory "<Path to files>" -SenderEmail "<Sender email>" -SmtpServer "<SMTP-Server>" -SmtpPort <Port>
``` 
- ExcelFile: (Required) The path to the Excel file containing the recipient information.
- FileDirectory: (Required) The path to the directory in which the files to be attached are located.
- SenderEmail: (Required) The e-mail address of the sender.
- SmtpServer: (Required) The SMTP server via which the emails are sent.
- SmtpPort: (Required) The port for the SMTP server.
- UserName: (Optional) The user name for authentication. If not specified, it is set to SenderEmail.
- LogFile: (Optional) The path to the log file. If not specified, the log file is created in the current script directory.
- MailSubject: (Optional) The subject of the email. “MAIL-TEST” by default.
- MailText: (Optional) The text of the email. By default “Hello [name]!”.

**Example call**:
``` 
.\mail.ps1 -ExcelFile "C:\Users\<user>\Desktop\Names.xlsx" -FileDirectory "C:\Users\<user>\Desktop\Files" -SenderEmail "youremail@gmx.de" -SmtpServer "mail.gmx.net" -SmtpPort 587
```

## Disclaimer
You use it at your own risk. I take no responsibility for any damages or problems that may arise from the use of this mail script.
