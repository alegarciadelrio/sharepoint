[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name AzureAD
Install-Module -Name ExchangeOnlineManagement
Import-Module  AzureAD
Import-Module ActiveDirectory
$msgWelcome = "Script to disable an user account of AD and AAD Services..."
$msgInitProcessAzure = "Connecting to Azure AD..."
$msgInitProcessExchange = "Connecting to Exchange Online..."
$msgInitProcessActiveDirectory = "Connecting to Active Directory..."
$msgInitProcessSyncAAD = "Syncing Azure AD..."
$msgInitProcessRemoveAAD = "Removing group membership of Azure AD..."
$msgInitProcessRenameAndMove = "Renaming and moving the AD's username..."
$msgFinish = "The process is finished..."
$name = ""
$surname = ""
$username = ""
$email = ""
$parameterRename = "z_archive"


$msgWelcome
$name = Read-Host 'Enter the name of the user (For example James)'
$surname = Read-Host 'Enter the surname of the user (For example Guo)'
$username = Read-Host 'Enter the username (For example James.Guo)'
$email = Read-Host 'Enter the email (For example james.guo@birchandwaite.com.au)'

$msgInitProcessAzure
Connect-AzureAD
Set-AzureADUser -ObjectID $email -AccountEnabled $false

$msgInitProcessExchange
Connect-ExchangeOnline
Set-Mailbox -Identity $email -Type Shared

$msgInitProcessActiveDirectory
Disable-ADAccount -Identity $username
$adGroupsOfUser = Get-ADPrincipalGroupMembership -Identity  $username | where {$_.Name -ne “Domain Users”}
# Removing group membership.
Remove-ADPrincipalGroupMembership -Identity  $username -MemberOf $adGroupsOfUser -Confirm:$false -verbose

$msgInitProcessSyncAAD
Invoke-Command MK-AZUREAD-W19V -Credential BWSRVR.CORP\Administrator { Start-ADSyncSyncCycle -PolicyType Delta }
Start-Sleep -Seconds 60

$msgInitProcessRemoveAAD
$aadUser = Get-AzureADUser -ObjectId $email
$aadGroupsOfUser = Get-AzureADUserMembership -ObjectId $email | where {$_.DisplayName -ne “All Users”}
foreach($group in $aadGroupsOfUser.ObjectId){
    Remove-AzureADGroupMember -ObjectId $group -MemberId $aadUser.ObjectId
}

$msgInitProcessRenameAndMove
Set-ADUser -Identity $username -DisplayName "z_archive $name $surname" -EmailAddress "z_archive.$email" -UserPrincipalName "z_archive.$username"  -SamAccountName "z_archive.$name"
Get-ADUser "z_archive.$name"| Move-ADObject -TargetPath 'OU=Archived Mailboxes,OU=Users,OU=theOU,DC=yourdomain,DC=CORP'
Get-ADUser "z_archive.$name"| Rename-ADObject -NewName "z_archive $name $surname"

$msgFinish
