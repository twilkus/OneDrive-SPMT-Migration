$userName = 
$passWord = Get-Content C:\1drive_migration\EP.txt | ConvertTo-SecureString -Force
$Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $passWord

$arg = '-Executionpolicy Bypass -File "C:\1drive_migration\SPMT\HomeDrive-Finalize.ps1"'
Start-Process -FilePath "PowerShell.exe" -ArgumentList $arg -Credential $Creds
