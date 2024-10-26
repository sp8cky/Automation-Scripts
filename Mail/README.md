# Mail automation script
This PowerShell script automates the sending of emails with attachments based on a list of recipients stored in an Excel file. It checks whether the files to be sent exist and logs the success or failure of each email transmission in a log file.

## Table of contents


## Implemented features
- Reads the recipient information (name and e-mail address) from a specified Excel file.
- Sends e-mails with attached files from a specified directory.
- Logs the sending status in a log file.
- Supports user-defined SMTP server settings.

## Installation and usage
### Anforderungen
- **PowerShell**: This script requires PowerShell 5.1 or higher.

#### Excel Datei
The excel file need two columns (every entry one row):
- Name: The name of the recipient.
- Email: The recipient's email address.


Clone the repository and navigate into the directory

git clone https://github.com/sp8cky/Automation-Scripts/Mail && cd Automation-Scripts/Mail

Install dependencies
```
pip install -r requirements.txt
```



### Aufruf 
Change the rights (see here for more information)
```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
``` 

.\script.ps1 -ExcelFile "<Pfad zur Excel-Datei>" -FileDirectory "<Pfad zum Verzeichnis>" -SenderEmail "<Deine E-Mail>" -SmtpServer "<SMTP-Server>" -SmtpPort <Port>

- ExcelFile: (Required) The path to the Excel file containing the recipient information.
- FileDirectory: (Required) The path to the directory in which the files to be attached are located.
- SenderEmail: (Required) The e-mail address of the sender.
- SmtpServer: (Required) The SMTP server via which the emails are sent.
- SmtpPort: (Required) The port for the SMTP server.
- UserName: (Optional) The user name for authentication. If not specified, it is set to SenderEmail.
- LogFile: (Optional) The path to the log file. If not specified, the log file is created in the current script directory.
- MailSubject: (Optional) The subject of the email. “MAIL-TEST” by default.
- MailText: (Optional) The text of the email. By default “Hello [name]!”.

Example call:
.\script.ps1 -ExcelFile "C:\Users\julia\Desktop\Names.xlsx" -FileDirectory "C:\Users\julia\Desktop\Files" -SenderEmail "jh4112000@gmx.de" -SmtpServer "mail.gmx.net" -SmtpPort 587

## Disclaimer
You use it at your own risk. I take no responsibility for any damages or problems that may arise from the use of this mail script.