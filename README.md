# Outlook-Excel-PowerShell
A homophonous PowerShell script that mines your Outlook inbox for emails attachments of specified names and outputs/updates the results in an Excel file.


Usage: 

Run PowerShell as Administrator,
Set-ExecutionPolicy RemoteSigned //enables scripts on your system (http://stackoverflow.com/questions/4037939/powershell-says-execution-of-scripts-is-disabled-on-this-system) 

test account is IMAP configured

bibliography:
ps in outlook:
https://msdn.microsoft.com/en-us/magazine/dn189202.aspx
http://stackoverflow.com/questions/22077808/reading-file-attachments-from-outlook-from-a-shared-mailbox
ps save to excel: 
http://sqlmag.com/powershell/update-excel-spreadsheets-powershell


mail.dll (for .NET) http://www.limilabs.com/mail
other .NET email: http://stackoverflow.com/questions/1159938/programmatically-open-an-email-from-a-pop3-and-extract-an-attachment