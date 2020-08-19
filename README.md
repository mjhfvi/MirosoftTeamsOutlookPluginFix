# Mirosoft Teams Outlook Plugin Fix
if your microsoft Teams Plugin to Outlook is not registering this is a PowerShell tool to check the regestry keys and add the keys you need



1. It will close outlook and teams,
2. Check if the system have the registry keys, if they have it will skip if they don’t have it will update the keys.
3. It will grab the folder version number for the plugin (%LOCALAPPDATA%\Microsoft\TeamsMeetingAddin) and add it in the registry key
4. check if the new folder have a ,dead file and remove it
5.And at the end it will open teams then Outlook..


if you need a manual command to fix it, run CMD as administrator.
C:\WINDOWS\system32\regsvr32.exe /s /n /i:user "%LOCALAPPDATA%\Microsoft\TeamsMeetingAddin\*GetVersionNumberFromDirectory\x64\Microsoft.Teams.AddinLoader.dll"
%LOCALAPPDATA%\Microsoft\Teams\Update.exe --processStart "Teams.exe" --process-start-args "--system-initiated"


*GetVersionNumberFromDirectory, replace this value with the number of the local diractory

#Tested on office 365 64Bit 
