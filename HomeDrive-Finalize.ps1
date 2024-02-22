###---Create input box
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Move (H) files to Moved2OneDrive'
$form.Size = New-Object System.Drawing.Size(500,200)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

$label = New-Object System.Windows.Forms.Label
$label.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
$label.Location = New-Object System.Drawing.Point(10,10)
$label.Size = New-Object System.Drawing.Size(480,40)
$label.Text = 'Please enter the email address of the target user:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Regular)
$textBox.Location = New-Object System.Drawing.Point(10,50)
$textBox.Size = New-Object System.Drawing.Size(430,20)
$textBox.Multiline = $false
$textBox.Text = "@williamblair.com"
$form.Controls.Add($textBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
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

###---Convert email handle to full email address and then retreive HomeDir value
$endUser = Get-ADUser -Filter {UserPrincipalName -eq $emailID } -Properties HomeDirectory
If ( $endUser -ne $null){
$homeDrive = $endUser.HomeDirectory
}
Else
{
Write-Host "Error. Please enter valid email handle."
Exit
}

###---Create the destination folder "Moved2OneDrive" and local log folder
If ( (Test-Path "$homeDrive\Moved2OneDrive") -eq $false)
{
    New-Item -Path $homeDrive -Name "Moved2OneDrive" -ItemType directory
}

If ( (Test-Path "C:\Blair\Logs\OneDrive_Migration") -eq $false)
{
    New-Item -Path "C:\Blair\Logs" -Name "OneDrive_Migration" -ItemType directory
}

###---Move content in H: drive to subfolder "Moved2OneDrive" w. log file
$Destination = "$homeDrive\Moved2OneDrive"
$timestamp = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
robocopy $homeDrive $destination /mir /move /sec /secfix /mt:16 /tbd /r:2 /w:5 /tee /np /v /log:"C:\Blair\Logs\OneDrive_Migration\"$emailID"_H_Finalization_"$timestamp".txt" /xd $Destination


###---Changes the ACL on the "Moved2OneDrive" folder by removing all access, disabling inheritance, setting ReadandExecute only access for the end-user and finally restoring system administrator access
$samAccount = (Get-ADUser -Filter {UserPrincipalName -eq $emailID } -Properties HomeDirectory).SamAccountName
$userAccount = 'BLAIRNET\' + $samAccount
$acl = Get-Acl $homeDrive
$acl.SetAccessRuleProtection($True, $False)
$acl.Access | % { $acl.RemoveAccessRule($_) }
$ruleRead = New-Object System.Security.AccessControl.FileSystemAccessRule(
"$userAccount",
"ReadAndExecute",
"ObjectInherit, ContainerInherit",
"None",
"Allow"
)
$acl.AddAccessRule($ruleRead)
(Get-Item "$homeDrive\Moved2OneDrive").SetAccessControl($acl)

$users = 
foreach ($user in $users) {
$ruleWrite = New-Object System.Security.AccessControl.FileSystemAccessRule(
"${user}",
"FullControl",
"ObjectInherit, ContainerInherit",
"None",
"Allow"
)
$acl.AddAccessRule($ruleWrite)
(Get-Item "$homeDrive\Moved2OneDrive").SetAccessControl($acl)
}
