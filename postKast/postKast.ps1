#postKast developed by wiregh0st

Write-Host "#####################################################################################
#                                                                                   #
#                Welcome to postKast developed by wiregh0st                         #
#                                                                                   #
#                                                                                   #        
#            You must change the paths at the top of postKast.ps1                   #
#            script if the file locations change at any time. If you do             #
#            not change the paths, the script will not execute properly!            #
#                                                                                   #
#####################################################################################"

Write-Host "`n"
Write-Host "Sleeping 10s before starting..."

$namespaceName = "root\cimv2\mdm\dmmap"
$className = "MDM_EnterpriseModernAppManagement_AppManagement01"
$wmiObj = Get-WmiObject -Namespace $namespaceName -Class $className
$result = $wmiObj.UpdateScanMethod()
Write-Host "Sleeping 30s to wait for Windows Store Updates..."
Start-Sleep -s 30

#declare variables for file paths. To call executables properly in powershell, use "&" before the variable name.
$xmlPath="J:\postKast\CustomDefaultAssoc.xml"
$defaultBrowserEXE="J:\postKast\SetDefaultBrowser.exe"
$syspinEXE="J:\postKast\syspin.exe"
$programDataPath="C:\ProgramData\Microsoft\Windows\Start Menu\Programs\"

#create directory in C: to store edge shortcut. Also creates variable that points to that path.
mkdir C:\system\.special | Out-Null
$sysPath = "C:\system\.special"

#function unpins apps from start. if use -pin, will pin app to start instead.
function Pin-App { param(
[string]$appname,
[switch]$unpin
)
try{
if ($unpin.IsPresent){
((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'From "Start" UnPin|Unpin from Start'} | %{$_.DoIt()}
return "App '$appname' unpinned from Start"
}else{
((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'To "Start" Pin|Pin to Start'} | %{$_.DoIt()}
return "App '$appname' pinned to Start"
}
}catch{
Write-Error "Error Pinning/Unpinning App! (App-Name correct?)"
}
}

Pin-App "Mail" -unpin
Pin-App "Microsoft Store" -unpin
Pin-App "Calendar" -unpin
Pin-App "Microsoft Edge" -unpin
Pin-App "Photos" -unpin
Pin-App "Cortana" -unpin
Pin-App "Weather" -unpin
Pin-App "Groove Music" -unpin
Pin-App "Xbox Console Companion" -unpin
Pin-App "movies & tv" -unpin
Pin-App "Office" -unpin
Pin-App "onenote" -unpin
Pin-App "Sticky Notes" -unpin
Pin-App "Network Speed Test" -unpin
Pin-App "Skype" -unpin
Pin-App "Microsoft Remote Desktop" -unpin
Pin-App "Remote Desktop" -unpin
Pin-App "Microsoft Whiteboard" -unpin
Pin-App "Microsoft To-Do" -unpin
Pin-App "Sway" -unpin
Pin-App "Calculator" -unpin
Pin-App "Office Lens" -unpin
Pin-App "Maps" -unpin
Pin-App "Alarms & Clock" -unpin
Pin-App "Voice Recorder" -unpin
Pin-App "Xbox" -unpin
Pin-App "Microsoft News" -unpin
Pin-App "News" -unpin

#create object to create microsoft edge shortcut. Allows pinning to taskbar. WILL NOT WORK OTHERWISE.
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut(($sysPath, "Edge.lnk" -join "\"))
$Shortcut.TargetPath = "shell:Appsfolder\Microsoft.MicrosoftEdge_8wekyb3d8bbwe!microsoftedge"
$Shortcut.IconLocation = "%windir%\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe"
$Shortcut.Save()

#create variable that paths to edge shortcut. also hide shortcut
$edgePath = $sysPath, "Edge.lnk" -join "\"
Get-ChildItem -path $sysPath -Recurse -Force | foreach {$_.attributes = "Hidden"}

#add paths to registry to disable protected view in office applications
$wshell = New-Object -ComObject wscript.shell;
Start-Process WINWORD
Start-Sleep -s 10
$wshell.SendKeys('{UP}')
Start-Sleep -s 3
$wshell.SendKeys('{ENTER}')
Start-Sleep -s 2
Stop-Process -name WINWORD -Force

New-Item -Path "HKCU:\software\microsoft\office\16.0\Word\Security\" -Name ProtectedView
New-Item -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\Security\" -Name ProtectedView
New-Item -Path "HKCU:\software\microsoft\office\16.0\Excel\Security\" -Name ProtectedView
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\Word\Security\ProtectedView" -Name 'DisableAttachmentsInPV' -value 1
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\Word\Security\ProtectedView" -Name 'DisableInternetFilesInPV' -value 1
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\Word\Security\ProtectedView" -Name 'DisableUnsafeLocationsInPV' -value 1
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\Security\ProtectedView" -Name 'DisableAttachmentsInPV' -value 1
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\Security\ProtectedView" -Name 'DisableInternetFilesInPV' -value 1
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\Security\ProtectedView" -Name 'DisableUnsafeLocationsInPV' -value 1
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\Excel\Security\ProtectedView" -Name 'DisableAttachmentsInPV' -value 1
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\Excel\Security\ProtectedView" -Name 'DisableInternetFilesInPV' -value 1
Set-Itemproperty -Path "HKCU:\software\microsoft\office\16.0\Excel\Security\ProtectedView" -Name 'DisableUnsafeLocationsInPV' -value 1



#import default file types .xml file. Also imports default email application (outlook).
Dism /online /import-defaultappassociations:$xmlPath

#select default browser
Write-Host "***Please select user preferred default browser***"
Write-Host "1: Mozilla Firefox"
Write-Host "2: Google Chrome"
Write-Host "3: Internet Explorer"
Write-Host "4: Microsoft Edge"

[Int]$selection = 0

$selection = Read-Host

#while loop for selection out of range
while(($selection -le 0) -or ($selection -ge 5)){
    Write-Host "ERROR: Value for selection out of range. Current value is: $selection."
	Write-Host "`n"
	Write-Host "***Please select user preferred default browser***"
	Write-Host "1: Mozilla Firefox"
	Write-Host "2: Google Chrome"
	Write-Host "3: Internet Explorer"
	Write-Host "4: Microsoft Edge"
    $selection = Read-Host
    }



switch($selection){
 
    1 {&$defaultBrowserEXE HKLM "FIREFOX.EXE"}
    2 {&$defaultBrowserEXE HKLM "Google Chrome"}
    3 {&$defaultBrowserEXE HKLM "IEXPLORE.EXE"}
    4 {&$defaultBrowserEXE HKLM "Edge"}
}

New-Item -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\" -Name Start
Set-Itemproperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start" -Name 'AllowPinnedFolderDownloads' -value 1
Set-Itemproperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start" -Name 'AllowPinnedFolderDownloads_ProviderSet' -value 1
Set-Itemproperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start" -Name 'AllowPinnedFolderFileExplorer' -value 1
Set-Itemproperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start" -Name 'AllowPinnedFolderFileExplorer_ProviderSet' -value 1

#delete registry key for taskbar icons then add taskbar browser icon via syspin.
Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" -Recurse

switch($selection){
    
    1 {&$syspinEXE "C:\Program Files (x86)\Mozilla Firefox\firefox.exe" 5386}
    2 {&$syspinEXE "C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe" 5386}
    3 {&$syspinEXE "C:\Program Files (x86)\Internet Explorer\iexplore.exe" 5386}
    4 {&$syspinEXE $edgePath 5386}
}

#add remaining taskbar icons.
&$syspinEXE ($programDataPath,"Word 2016.lnk" -join "\") 5386
&$syspinEXE ($programDataPath,"Outlook 2016.lnk" -join "\") 5386
&$syspinEXE ($programDataPath,"Excel 2016.lnk" -join "\") 5386
&$syspinEXE ($programDataPath,"PowerPoint 2016.lnk" -join "\") 5386


Write-Host "Post cast finished. Restarting explorer... "

Stop-Process -name explorer

Write-Host "Sleeping 10 seconds and then closing... "

Start-Sleep -s 10

exit