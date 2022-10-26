$destinationuser = Read-Host "Enter destination user's email"
$site = Read-Host "Enter the name of the site"
$globaladmin = Read-Host "Enter the username of your Global Admin account"
$credentials = Get-Credential -Credential $globaladmin
Connect-MsolService -Credential $credentials

$InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}

#connects and adds User as admin on Onedrive
$SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"
  

  
$departingOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0]).sharepoint.com/sites/$site"
Write-Host "`nConnecting to SharePoint Online" -ForegroundColor Green
Connect-SPOService -Url $SharePointAdminURL -Credential $credentials

Write-Host "`nAdding $destinationUserUnderscore as site collection admin on $departingOneDriveSite" -ForegroundColor Green
Set-SPOUser -Site $departingOneDriveSite -LoginName $destinationuser -IsSiteCollectionAdmin $true