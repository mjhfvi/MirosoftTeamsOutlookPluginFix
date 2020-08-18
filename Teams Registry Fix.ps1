#####################################################################
#  check if all the key are in the system
#
#  click of exit outlook and teams
#  run the registry fix
#
#  start outlook and teams
#
#  pop "We are done, please go to outlook and test the plugin"
#
######################################################################

### Close Outlook and Teams
$ProcessOutlook = Get-Process -Name Outlook -ErrorAction SilentlyContinue
if ($ProcessOutlook) {$ProcessOutlook.CloseMainWindow()
Sleep 2
if (!$ProcessOutlook.HasExited) {$ProcessOutlook | Stop-Process -Force}}

$ProcessTeams = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
if ($ProcessTeams) {$ProcessTeams | Stop-Process -Force}


### Get the Version from local folder in LOCALAPPDATA\microsoft\teamsmeetingaddin\
$TeamsFolderVersion = Get-ChildItem -Path $env:LOCALAPPDATA\microsoft\teamsmeetingaddin\ -Directory -Force -ErrorAction SilentlyContinue | foreach Name | Select-Object -last 1
$TeamsFolderVersion
$FolderPathEnd = "\x64\Microsoft.Teams.AddinLoader.dll"


### Test the Registry Keys
$TestKey01 = Test-Path "HKCU:\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}"
$TestKey01
IF($TestKey01 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\" -Name "{19A6E644-14E6-4A60-B8D7-DD20610A871D}")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}" -Name "(Default)" -Value "FastConnect Class" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\" -Name InprocServer32)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\InprocServer32" -Name "(Default)" -Value $Env:localappdata\microsoft\teamsmeetingaddin\$TeamsFolderVersion\$FolderPathEnd -Force | Out-Null)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\InprocServer32" -Name "ThreadingModel" -Value "Apartment" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\" -Name ProgID)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\ProgID" -Name "(Default)" -Value "TeamsAddin.FastConnect.1" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\" -Name Programmable)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\Programmable" -Name "(Default)" -Value  -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\" -Name TypeLib)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\TypeLib" -Name "(Default)" -Value "{C0529B10-073A-4754-9BB0-72325D80D122}" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\" -Name Version)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\Version" -Name "(Default)" -Value "1.0" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\" -Name VersionIndependentProgID)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\VersionIndependentProgID" -Name "(Default)" -Value "TeamsAddin.FastConnect" -Force | Out-Null)
}

$TestKey02 = Test-Path "HKCU:\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}"
$TestKey02
IF($TestKey02 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\" -Name "{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}" -Name "(Default)" -Value "Connect Class" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\" -Name InprocServer32)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\InprocServer32" -Name "(Default)" -Value $Env:localappdata\microsoft\teamsmeetingaddin\$TeamsFolderVersion\$FolderPathEnd -Force | Out-Null)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\InprocServer32" -Name "ThreadingModel" -Value "Apartment" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\" -Name ProgID)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\ProgID" -Name "(Default)" -Value "TeamsAddin.Connect.1" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\" -Name Programmable)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\Programmable" -Name "(Default)" -Value  -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\" -Name TypeLib)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\TypeLib" -Name "(Default)" -Value "{C0529B10-073A-4754-9BB0-72325D80D122}" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\" -Name Version)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\Version" -Name "(Default)" -Value "1.0" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\" -Name VersionIndependentProgID)
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\VersionIndependentProgID" -Name "(Default)" -Value "TeamsAddin.Connect" -Force | Out-Null)
}

$TestKey03 = Test-Path "HKCU:\Software\Classes\TeamsAddin.FastConnect"
$TestKey03
IF($TestKey03 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Classes\" -Name "TeamsAddin.FastConnect")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\TeamsAddin.FastConnect" -Name "(Default)" -Value "FastConnect Class" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\TeamsAddin.FastConnect" -Name "CurVer")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\TeamsAddin.FastConnect" -Name "(Default)" -Value "TeamsAddin.FastConnect.1" -Force | Out-Null)
}

$TestKey04 = Test-Path "HKCU:\Software\Classes\TeamsAddin.FastConnect.1"
$TestKey04
IF($TestKey04 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Classes\" -Name "TeamsAddin.FastConnect.1")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\TeamsAddin.FastConnect.1" -Name "(Default)" -Value "FastConnect Class" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\TeamsAddin.FastConnect.1" -Name "CLSID")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\TeamsAddin.FastConnect.1" -Name "(Default)" -Value "{19A6E644-14E6-4A60-B8D7-DD20610A871D}" -Force | Out-Null)
}

$TestKey05 = Test-Path "HKCU:\Software\Classes\TeamsAddin.Connect"
$TestKey05
IF($TestKey05 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Classes\" -Name "TeamsAddin.Connect")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\TeamsAddin.Connect" -Name "(Default)" -Value "Connect Class" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\TeamsAddin.Connect" -Name "CurVer")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\TeamsAddin.Connect" -Name "(Default)" -Value "TeamsAddin.Connect.1" -Force | Out-Null)
}

$TestKey06 = Test-Path "HKCU:\Software\Classes\TeamsAddin.Connect.1"
$TestKey06
IF($TestKey06 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Classes\TeamsAddin.Connect" -Name "TeamsAddin.Connect.1")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\TeamsAddin.Connect.1" -Name "(Default)" -Value "Connect Class" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\Software\Classes\TeamsAddin.Connect.1" -Name "CLSID")
(New-ItemProperty -Path Registry::"HKCU\Software\Classes\TeamsAddin.Connect.1" -Name "(Default)" -Value "{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}" -Force | Out-Null)
}

$TestKey07 = Test-Path "HKCU:\Software\Classes\Wow6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}"
$TestKey07
IF($TestKey07 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "{19A6E644-14E6-4A60-B8D7-DD20610A871D}")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}" -Name "(Default)" -Value "FastConnect Class" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}" -Name "InprocServer32")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\InprocServer32" -Name "(Default)" -Value $Env:localappdata\microsoft\teamsmeetingaddin\$TeamsFolderVersion\$FolderPathEnd -Force | Out-Null)
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\InprocServer32" -Name "ThreadingModel" -Value "Apartment" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "ProgID")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\ProgID" -Name "(Default)" -Value "TeamsAddin.FastConnect.1" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "Programmable")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\Programmable" -Name "(Default)" -Value  -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "TypeLib")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\TypeLib" -Name "(Default)" -Value "{C0529B10-073A-4754-9BB0-72325D80D122}" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "Version")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\Version" -Name "(Default)" -Value "1.0" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "VersionIndependentProgID")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\VersionIndependentProgID" -Name "(Default)" -Value "TeamsAddin.FastConnect" -Force | Out-Null)
}

$TestKey08 = Test-Path "HKCU:\Software\Classes\Wow6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}"
$TestKey08
IF($TestKey08 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}" -Name "(Default)" -Value "Connect Class" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}" -Name "InprocServer32")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\InprocServer32" -Name "(Default)" -Value $Env:localappdata\microsoft\teamsmeetingaddin\$TeamsFolderVersion\$FolderPathEnd -Force | Out-Null)
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\InprocServer32" -Name "ThreadingModel" -Value "Apartment" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "ProgID")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\ProgID" -Name "(Default)" -Value "TeamsAddin.Connect.1" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "Programmable")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\Programmable" -Name "(Default)" -Value  -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "TypeLib")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\TypeLib" -Name "(Default)" -Value "{C0529B10-073A-4754-9BB0-72325D80D122}" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "Version")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\Version" -Name "(Default)" -Value "1.0" -Force | Out-Null)
(New-Item -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\" -Name "VersionIndependentProgID")
(New-ItemProperty -Path Registry::"HKCU\SOFTWARE\Classes\WOW6432Node\CLSID\{CB965DF1-B8EA-49C7-BDAD-5457FDC1BF92}\VersionIndependentProgID" -Name "(Default)" -Value "TeamsAddin.Connect" -Force | Out-Null)
}

$TestKey09 = Test-Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList"
$TestKey09
IF($TestKey09 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Microsoft\Office\15.0\Outlook\Resiliency\" -Name "DoNotDisableAddinList")
(New-ItemProperty -Path Registry::"HKCU\Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList" -Name "TeamsAddin.Connect" -Value "1" -PropertyType DWORD -Force | Out-Null)
}

$TestKey10 = Test-Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"
$TestKey10
IF($TestKey10 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\" -Name "DoNotDisableAddinList")
(New-ItemProperty -Path Registry::"HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" -Name "TeamsAddin.Connect" -Value "1" -PropertyType DWORD -Force | Out-Null)
}

$TestKey11 = Test-Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
$TestKey11
IF($TestKey11 -eq "True") {'The Key Exist'} Else {
(New-Item -Path Registry::"HKCU\Software\Microsoft\Office\Outlook\Addins\" -Name "TeamsAddin.FastConnect")
(New-ItemProperty -Path Registry::"HKCU\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Name "Description" -Value "Microsoft Teams Meeting Add-in for Microsoft Office" -Force | Out-Null)
(New-ItemProperty -Path Registry::"HKCU\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Name "FriendlyName" -Value "Microsoft Teams Meeting Add-in for Microsoft Office" -Force | Out-Null)
(New-ItemProperty -Path Registry::"HKCU\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Name "LoadBehavior" -Value "3" -PropertyType DWORD -Force | Out-Null)
}

### Remove .dead file
$FileDead = "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin\$TeamsFolderVersion\.dead"
if (Test-Path $FileDead)
{
Remove-Item $FileDead
}


### Close Outlook and Teams
Start-Sleep -s 5
Start-Process "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe"

Start-Sleep -s 10
Start-Process "Outlook"
