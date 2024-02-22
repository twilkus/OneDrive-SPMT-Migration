<#
Vendor = 'Microsoft'
Application = 'Sharepoint Migration Tool'
Script Date = '2/4/2022'
Script Version  = '1.8
Script Author(s) = 'Tom Wilkus & Mile Siriski"

Change Log:
2/14/2022 - added code for input box which discovers and populates 'source' and 'destination' variables based on user's email address
2/14/2022 - configured skipped file extensions to: one, onetoc2, exe, msi, bin, pst
2/14/2022 - set working folder for logs and reports to '\\S-ONEDRIVE0-V\C$\1drive_migration\_Reports'
2/17/2022 - adding Get-DirectoryTreeSize cmdlet
2/23/2022 - added hashed password for $SPOCredential from text file
2/25/2022 - changed working folder back to relative target 'C:\1drive_migration\Reports' due to remote execution issue
2/28/2022 - implemented & configured custom email notifications
3/07/2022 - enabled "MigrateWithoutRootFolder" and disabled  "ParametersValidationOnly" effectively making this a PROD script
3/08/2022 - added check for OneDrive account associated with target user with error email
3/10/2022 - changed notification email to "onedrivemigrationstatus@williamblair.com"
3/30/2022 - add user email to body of start and end migration emails
4/07/2022 - included OneNote files into migration
4/08/2022 - revision 2.0
5/06/2022 - automate adding user to OneDrive folder redirection AD group
5/10/2022 - fix automate H: drive data consolidation into "H:\Moved2OneDrive" folder


TO-DO:
- move reports folder to S: drive
#>

###---Set target email for notifications and VDI hostname
$notificationEmail = 
$Hostname = $ENV:COMPUTERNAME

###---Set the script path to current location
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition

###---Dot source Migration_Functions.ps1
. "C:\1drive_migration\SPMT\Launch_Batch_Functions.ps1"

###---Import DirectoryTreeSize module from SPMT folder for this PS session
$Env:PSModulePath = $Env:PSModulePath+";C:\1drive_migration\SPMT"
Import-Module DirectoryTreeSize

###---Import SPMT migration module
Import-Module Microsoft.SharePoint.MigrationTool.PowerShell

#region Form

###---Create input box for email address
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(490,200)
$form.Text = '  ~Migrate to OneDrive~'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

$label = New-Object System.Windows.Forms.Label
$label.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
$label.Location = New-Object System.Drawing.Point(10,10)
$label.Size = New-Object System.Drawing.Size(490,30)
$label.Text = 'Please enter the email address of the target user:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Regular)
$textBox.Location = New-Object System.Drawing.Point(12,55)
$textBox.Size = New-Object System.Drawing.Size(435,20)
$textBox.Multiline = $false
$textBox.Text = "@williamblair.com"
$form.Controls.Add($textBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
$okButton.Location = New-Object System.Drawing.Point(12,110)
$okButton.Size = New-Object System.Drawing.Size(80,30)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
$cancelButton.Location = New-Object System.Drawing.Point(100,110)
$cancelButton.Size = New-Object System.Drawing.Size(80,30)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

If ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $emailID = $textBox.Text
    $emailID | Out-Null
}
Else{
    Break
}

#endregion Form

###---Convert email address to HomeDir value as network share source folder
$homeDrive = Get-ADUser -Filter {UserPrincipalName -eq $emailID} -Properties HomeDirectory
If ($homeDrive.HomeDirectory -ne $null){
    $FileshareSource = $homeDrive.HomeDirectory
}

###---Report if H: drive was not found for the user
Else{
    Write-Host "***Error: Unable to find H: Drive for $emailID; migration for user failed. Check user's email address and AD then restart migration." "`n" -ForegroundColor Red
    Send-Email -To "$notificationEmail" -Subject "ERROR: Missing migration information" -Body "Unable to find H: Drive for $emailID; migration for user failed. Check user's email address and AD then restart migration."
    Break
}

###---Retreive OneDrive URL based on email address and select a 'Personal Space' only
$userName = 
$Password = Get-Content C:\1drive_migration\EPP.txt | ConvertTo-SecureString -Force
$SPOCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $Password
Connect-SPOService -Url https://williamblair-admin.sharepoint.com -Credential $SPOCredential
$sODsite = (Get-SPOSite -IncludePersonalSite $true -Filter "owner -eq $emailID").url

###---Check and report if OneDrive account was not found for the user or continue
If($sODsite -eq $null){
    Write-Host "***Error: Unable to find OneDrive account for $emailID; migration for user failed. Reach out to Endpoint Management." "`n" -ForegroundColor Red
    Send-Email -To "$notificationEmail" -Subject "ERROR: Missing migration information" -Body "Unable to find OneDrive account for $emailID; migration for user failed. Reach out to Endpoint Management."
    Break
}

If ($sODsite[1])
{
    ForEach ($sODsubsite in $sODsite)
    {
        If ($sODsubsite.Contains("personal"))
        {
            $sODsite = $sODsubsite
        }
    }
}

###---Register the SPMT session
Register-SPMTMigration -SPOCredential $SPOCredential -Force -WorkingFolder 'C:\1drive_migration\Reports' -SkipFilesWithExtension exe,msi,bin,pst -PreserveUserPermissionsForFileShare $true -MigrateWithoutRootFolder

###---Load SPMT session information and write to log file
$Session = Get-SPMTMigration
$CurrentID = $Session.Id
$CurrentReports = $Session.ReportFolderPath
$CurrentReportsSub = Split-Path $CurrentReports

###---Send email for batch job start
$Timestamp = Get-Date -Format "MM-dd-yyyy; HH:mm:ss"
Send-Email -To $notificationEmail -Subject "INFO: OneDrive Migration User Start" -Body "Migration:`t$CurrentID `nUser:`t`t$emailID `nStart Time:`t$timestamp `nHost VDI:`t$Hostname `nReports:`t$CurrentReportsSub `n`n***This mailbox is not monitored. Do not reply.***"

Start-Sleep 5

###---Add FileShare migration task to session
Add-SPMTTask -FileShareSource $FileshareSource -TargetSiteUrl $sODsite -TargetList "Documents"

###---Start migration in the Powershell console
Start-SPMTMigration -OutVariable Out

###---Add user to folder redirction AD group
$group = 
$userSamName = (Get-ADUser -Properties samAccountName -Filter { mail -like $emailID }).samAccountName
Add-ADGroupMember -Identity $group -Members $userSamName

###---Home drive finalization - move H: contents to H:\Moved2OneDrive
###---Create the destination folder "Moved2OneDrive"
If ((Test-Path "$FileshareSource\Moved2OneDrive") -eq $false)
{
	New-Item -Path $FileshareSource -Name "Moved2OneDrive" -ItemType directory
}
	
###---Create robocopy log folder if it does not exist
If ((Test-Path "C:\1drive_migration\Reports\Robocopy") -eq $false)
{
	New-Item -Path "C:\1drive_migration\Reports\" -Name "Robocopy" -ItemType directory
}
	
###---Move content in H: drive to subfolder "Moved2OneDrive" w. log file
$Destination = "$FileshareSource\Moved2OneDrive"
$timestamp = Get-Date -Format "yyyy-MM-dd"
$logfile = ("C:\1drive_migration\Reports\Robocopy\" + $emailID + "_H_Finalization_" + $timestamp + ".txt")
$source = "$FileshareSource"
$target = "$Destination"
$robocopyArgs = @($source, $target, "/move", "/mir", "/sec", "/secfix", "/mt:16", "/tbd", "/r:2", "/w:5", "/tee", "/np", "/v", "/log:$logfile", "/xd $Destination")
start-process robocopy $robocopyArgs -wait

###---Send email for batch job end
$Timestamp = Get-Date -Format "MM-dd-yyyy; HH:mm:ss"
Send-Email -To "$notificationEmail" -Subject "INFO: OneDrive Migration User Finish" -Body "Migration:`t$CurrentID `nUser:`t`t$emailID `nStart Time:`t$timestamp `nHost VDI:`t$Hostname `nReports:`t$CurrentReportsSub `n`n***This mailbox is not monitored. Do not reply.***"

###---Exit script and close migration (due to Stop-SPMTMigration not working)
Start-Sleep 5
Stop-Process -Id $PID