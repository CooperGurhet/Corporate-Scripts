#this code is a mess of variable names because I snagged it from other code I wrote

$destinationuser = Read-Host "Enter destination user's email"
$site = Read-Host "Enter the users email"
$globaladmin = Read-Host "Enter the username of your Global Admin account"
$credentials = Get-Credential -Credential $globaladmin
Connect-MsolService -Credential $credentials

$InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}

#connects and adds User as admin on Onedrive
$SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"

$siteUnderscore = $site -replace "[^a-zA-Z]", "_"

  
$departingOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$siteUnderscore"
Write-Host "`nConnecting to SharePoint Online" -ForegroundColor Green
Connect-SPOService -Url $SharePointAdminURL -Credential $credentials

Write-Host "`nAdding $destinationUserUnderscore as site collection admin on $departingOneDriveSite" -ForegroundColor Green
Set-SPOUser -Site $departingOneDriveSite -LoginName $destinationuser -IsSiteCollectionAdmin $true