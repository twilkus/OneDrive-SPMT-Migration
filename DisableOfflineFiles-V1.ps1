<#
Vendor = 'WB Endpoint Management'
Application = 'Disable Offline Files'
Script Date = '3/21/2022'
Script Version  = '1.0
Script Author = 'Tom Wilkus'

Description: Script that will disable Offline Files without resolving conflicts. This script will take ownership of the C:\Windows\CSC folder as the user, make a backup of the CSC folder, and disable Offline Files.
This script utilizes the Microsoft SubInACL utility to take ownership of the CSC folder. 

#>
#Functions

function Get-ScriptDirectory{
if($psISE)
    {
        Split-Path $psISE.CurrentFile.FullPath
    }
    else
    {
        $Global:PSScriptRoot
    }
}

function Get-LoggedInUser{
    $ExplorerProcess = gwmi win32_process | where name -Match explorer

    if($ExplorerProcess.getowner().user.count -gt 1){
        return $ExplorerProcess.getowner().user[0]
    }

    else{
        return $ExplorerProcess.getowner().user
    }
}


#Variables

$scriptPath = Get-ScriptDirectory
$CSCPath = "HKLM:SYSTEM\CurrentControlSet\Services\CSC"
$CSCServicePath = "HKLM:SYSTEM\CurrentControlSet\Services\CscService"
$CSCFolder = "C:\Windows\CSC"
$CSCBackupFolder = "C:\Blair\OfflineFilesBackup"
$userName = Get-LoggedInUser

#Run SubInACL utility to take ownership of CSC folder and subfolders as logged in user
$subinaclArgs1 = @("/errorlog=c:\blair\subinaclCSC_err.txt", "/outputlog=c:\blair\subinaclCSC.txt", "/file C:\Windows\CSC\", "/grant=BLAIRNET\$userName=F")
$subinaclArgs2 = @("/errorlog=c:\blair\subinaclCSCsubdirs_err.txt", "/outputlog=c:\blair\subinaclCSCsubdirs.txt", "/subdirectories C:\Windows\CSC\", "/grant=BLAIRNET\$userName=F")
Start-Process -FilePath "$scriptPath.\subinacl.exe" -ArgumentList $subinaclArgs1 -Wait
Start-Process -FilePath "$scriptPath.\subinacl.exe" -ArgumentList $subinaclArgs2 -Wait

#Disable Offline Files via registry. Offline Files is not fully disabled until a reboot takes place. 
Set-ItemProperty -Path $CSCPath -Name Start -Value 4
Set-ItemProperty -Path $CSCServicePath -Name Start -Value 4

#Create backup copy of CSC folder
if((Test-Path -Path $CSCBackupFolder) -eq $false)
{
    New-Item -Path $CSCBackupFolder -ItemType Directory -Force
}

Copy-Item -Path $CSCFolder -Destination $CSCBackupFolder -Recurse -Force


#Undo Folder Redirection and reset to default
<#
$UserProfileVar = '%USERPROFILE%'
$ShellFoldersRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
$UserShellFoldersRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

#Reset Shell Folders
$NewDesktopPath = $env:USERPROFILE + "\Desktop"
$NewFavoritesPath = $env:USERPROFILE + "\Favorites"
$NewDocumentsPath = $env:USERPROFILE + "\Documents"

Set-ItemProperty -path $ShellFoldersRegPath -name Desktop $NewDesktopPath
Set-ItemProperty -path $ShellFoldersRegPath -name Favorites $NewFavoritesPath
Set-ItemProperty -path $ShellFoldersRegPath -name Personal $NewDocumentsPath

#Reset User Shell Folders
$NewDesktopPath2 = $UserProfileVar + "\Desktop"
$NewFavoritesPath2 = $UserProfileVar + "\Favorites"
$NewDocumentsPath2 = $UserProfileVar + "\Documents"

Set-ItemProperty -path $UserShellFoldersRegPath -name Desktop $NewDesktopPath2
Set-ItemProperty -path $UserShellFoldersRegPath -name Favorites $NewFavoritesPath2
Set-ItemProperty -path $UserShellFoldersRegPath -name Personal $NewDocumentsPath2
#>