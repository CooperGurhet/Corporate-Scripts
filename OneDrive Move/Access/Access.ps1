$starttime = Get-Date
$departinguser = Read-Host "Enter departing user's email"
$destinationuser = Read-Host "Enter destination user's email"
$globaladmin = Read-Host "Enter the username of your Global Admin account"
$credentials = Get-Credential -Credential $globaladmin
Connect-MsolService -Credential $credentials

$InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}

#connects and adds User as admin on Onedrive
$SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"
  
$departingUserUnderscore = $departinguser -replace "[^a-zA-Z]", "_"
$destinationUserUnderscore = $destinationuser -replace "[^a-zA-Z]", "_"
  
$departingOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$departingUserUnderscore"
Write-Host "`nConnecting to SharePoint Online" -ForegroundColor Green
Connect-SPOService -Url $SharePointAdminURL -Credential $credentials

Write-Host "`nAdding $destinationUserUnderscore as site collection admin on $departingOneDriveSite" -ForegroundColor Green
Set-SPOUser -Site $departingOneDriveSite -LoginName $destinationuser -IsSiteCollectionAdmin $true

#connects and gives user mail permission
Connect-ExchangeOnline -UserPrincipalName $globaladmin

add-mailboxPermission -Identity $departinguser -User $destinationuser -accessRights FullAccess -inheritancetype All -Confirm:$false
Add-RecipientPermission -identity $departinguser -trustee $destinationuser -AccessRights "sendas" -Confirm:$false

Disconnect-ExchangeOnline -Confirm:$false

$email = "https://outlook.office.com/mail/$($departinguser)"

Write-Host "`nComplete!" -ForegroundColor Green
Write-Host "`nEmail: $($email)" -ForegroundColor Green
Write-Host "OneDrive: $($departingOneDriveSite)" -ForegroundColor Green
$endtime = Get-Date
Write-Host "started at: $starttime"
write-host "completed at: $endtime"