$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

#$computerName = 'WILKUST-V'
$computerName = 


Add-Content $scriptpath\DEV-Launch-Migration.log -Value "--------------------------------------------------------------------------------"
Add-Content $scriptpath\DEV-Launch-Migration.log -Value ("Timestamp: ".ToString() + (get-date -format "MM-dd-yyyy; HH:mm:ss"))

## PsExec v2.2 (from script path)
& $scriptPath\psexec.exe -accepteula -i \\$computerName PowerShell.exe -ExecutionPolicy Bypass -File "C:\1drive_migration\SPMT\SPMT-Migration-Batch-Launch.ps1" | Add-Content $scriptpath\Migration_launcher.log