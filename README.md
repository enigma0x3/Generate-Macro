#Coded by Matt Nelson (@enigma0x3)
<#
.SYNOPSIS
Standalone Powershell script that will generate a malicious Microsoft Office document with a specified payload and persistence method

.DESCRIPTION
This script will generate malicious Microsoft Excel Documents that contain VBA macros. This script will prompt you for your attacking IP 
(the one you will receive your shell at), the port you want your shell at, and the name of the document. From there, the script will then
display a menu of different attacks, all with different persistence methods. Once an attack is chosen, it will then prompt you for your payload type
(Only HTTP and HTTPS are supported).

When naming the document, don't include a file extension.

These attacks use Invoke-Shellcode, which was created by Matt Graeber. Follow him on Twitter --> @mattifestation


.Attack Types
Meterpreter Shell with Logon Persistence: This attack delivers a meterpreter shell and then persists in the registry 
by creating a hidden .vbs file in C:\Users\Public and then creates a registry key in HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load
that executes the .vbs file on login.

Meterpreter Shell with Powershell Profile Persistence: This attack requires the target user to have admin right but is quite creative. It will
deliver you a shell and then drop a malicious .vbs file in C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.vbs. Once dropped, it creates
an infected Powershell Profile file in C:\Windows\SysNative\WindowsPowerShell\v1.0\ and then creates a registry key in 
HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load that executes Powershell.exe on startup. Since the Powershell profile loads automatically when 
Powershell.exe is invoked, your code is executed automatically.

Meterpreter Shell with Microsoft Outlook Email Persistence: This attack will give you a shell and then download a malicious Powershell script in this location:
C:\Users\Public\. Once downloaded, it will insert your defined IP address, Port, Email address and Trigger word.
It will then create a malicious .vbs file and drop it in C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\. Once dropped, it creates
a registry key that executes it on login. When the Powershell script is executed, it monitors the user's Outlook Inbox for an email containing 
the email address you specified as well as the subject. When it sees the email, it will delete it and send you a shell.


.EXAMPLE
PS> ./Generate-Macro.ps1
Enter IP Address: 10.0.0.10
Enter Port Number: 1111
Enter the name of the document (Do not include a file extension): FinancialData

--------Select Attack---------
1. Meterpreter Shell with Logon Persistence
2. Meterpreter Shell with Powershell Profile Persistence (Requires user to be local admin)
3. Meterpreter Shell with Microsoft Outlook Email Persistence
------------------------------
Select Attack Number & Press Enter: 1

--------Select Payload---------
1. Meterpreter Reverse HTTPS
2. Meterpreter Reverse HTTP
------------------------------
Select Payload Number & Press Enter: 1
Saved to file C:\Users\Malware\Desktop\FinancialData.xls
PS>




#>
