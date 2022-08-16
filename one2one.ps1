$starttime = Get-Date
$departinguser = Read-Host "Enter departing user's email"
$destinationuser = Read-Host "Enter destination user's email"
$globaladmin = Read-Host "Enter the username of your Global Admin account"
$credentials = Get-Credential -Credential $globaladmin
Connect-MsolService -Credential $credentials
 
$InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}
  
$SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"
  
$departingUserUnderscore = $departinguser -replace "[^a-zA-Z]", "_"
$destinationUserUnderscore = $destinationuser -replace "[^a-zA-Z]", "_"
  
$departingOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$departingUserUnderscore"
$destinationOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$destinationUserUnderscore"
Write-Host "`nConnecting to SharePoint Online" -ForegroundColor Green
Connect-SPOService -Url $SharePointAdminURL -Credential $credentials
  
Write-Host "`nAdding $globaladmin as site collection admin on both OneDrive site collections" -ForegroundColor Green
# Set current admin as a Site Collection Admin on both OneDrive Site Collections
Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
Set-SPOUser -Site $destinationOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
  
Write-Host "`nConnecting to $departinguser's OneDrive via SharePoint Online PNP module" -ForegroundColor Green
  
Connect-PnPOnline -Url $departingOneDriveSite -Credentials $credentials
  
Write-Host "`nGetting display name of $departinguser" -ForegroundColor Green
# Get name of departing user to create folder name.
$departingOwner = Get-PnPSiteCollectionAdmin | Where-Object {$_.loginname -match $departinguser}
  
# If there's an issue retrieving the departing user's display name, set this one.
if ($departingOwner -contains $null) {
    $departingOwner = @{
        Title = "Departing User"
    }
}
  
# Define relative folder locations for OneDrive source and destination
$departingOneDrivePath = "/personal/$departingUserUnderscore/Documents"
$destinationOneDrivePath = "/personal/$destinationUserUnderscore/Documents/$($departingOwner.Title)'s Files"
$destinationOneDriveSiteRelativePath = "Documents/$($departingOwner.Title)'s Files"
  
Write-Host "`nGetting all items from $($departingOwner.Title)" -ForegroundColor Green
# Get all items from source OneDrive
$items = Get-PnPListItem -List Documents -PageSize 1000
  
$largeItems = $items | Where-Object {[long]$_.fieldvalues.SMTotalFileStreamSize -ge 261095424 -and $_.FileSystemObjectType -contains "File"}
if ($largeItems) {
    $largeexport = @()
    foreach ($item in $largeitems) {
        $largeexport += "$(Get-Date) - Size: $([math]::Round(($item.FieldValues.SMTotalFileStreamSize / 1MB),2)) MB Path: $($item.FieldValues.FileRef)"
        Write-Host "File too large to copy: $($item.FieldValues.FileRef)" -ForegroundColor DarkYellow
    }
    $largeexport | Out-file C:\temp\largefiles.txt -Append
    Write-Host "A list of files too large to be copied from $($departingOwner.Title) have been exported to C:\temp\LargeFiles.txt" -ForegroundColor Yellow
}
  
$rightSizeItems = $items | Where-Object {[long]$_.fieldvalues.SMTotalFileStreamSize -lt 261095424 -or $_.FileSystemObjectType -contains "Folder"}
  
Write-Host "`nConnecting to $destinationuser via SharePoint PNP PowerShell module" -ForegroundColor Green
Connect-PnPOnline -Url $destinationOneDriveSite -Credentials $credentials
  
Write-Host "`nFilter by folders" -ForegroundColor Green
# Filter by Folders to create directory structure
$folders = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "Folder"}
$i = 0
$total = $folders.Count
foreach ($folder in $folders) {
    $path = ('{0}{1}' -f $destinationOneDriveSiteRelativePath, $folder.fieldvalues.FileRef).Replace($departingOneDrivePath, '')
    $i++
    Write-Progress -Activity "Creating Directory Structure" -status "$i/$total" -PercentComplete (($i/$total)*100) -CurrentOperation "Creating folder in $path"
    $newfolder = Resolve-PnPFolder -SiteRelativePath $path
}
  

$files = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "File"}
$fileerrors = ""
$i = 0
$total = $files.Count
foreach ($file in $files) {
    $destpath = ("$destinationOneDrivePath$($file.fieldvalues.FileDirRef)").Replace($departingOneDrivePath, "")
    $i++
    Write-Progress -Activity "Creating Files" -status "$i/$total" -PercentComplete (($i/$total)*100) -CurrentOperation "Copying $($file.fieldvalues.FileLeafRef) to $destpath"
    $newfile = Copy-PnPFile -SourceUrl $file.fieldvalues.FileRef -TargetUrl $destpath -OverwriteIfAlreadyExists -Force -ErrorVariable errors -ErrorAction SilentlyContinue
    $fileerrors += $errors
}
$fileerrors | Out-File c:\temp\fileerrors.txt
  
# Remove Global Admin from Site Collection Admin role for both users
Write-Host "`nRemoving $globaladmin from OneDrive site collections" -ForegroundColor Green
Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
Set-SPOUser -Site $destinationOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
Write-Host "`nComplete!" -ForegroundColor Green
$endtime = Get-Date
Write-Host "started at: $starttime"
write-host "completed at: $endtime"
