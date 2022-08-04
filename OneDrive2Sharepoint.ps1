$departinguser = Read-Host "Enter departing user's email"
$DestinationSite = Read-Host "Enter destination site Name"
$globaladmin = Read-Host "Enter the username of your Global Admin account"
$credentials = Get-Credential -Credential $globaladmin
Connect-MsolService -Credential $credentials
 
$InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}
  
$SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"
  
$departingUserUnderscore = $departinguser -replace "[^a-zA-Z]", "_"
$DestinationSiteFormatted = $DestinationSite -replace "[^a-zA-Z]", ""
  
$departingOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$departingUserUnderscore"
$DestinationSharepointSite = "https://$($InitialDomain.Name.Split(".")[0]).sharepoint.com/sites/$DestinationSiteFormatted"

Write-Host "`nConnecting to SharePoint Online" -ForegroundColor Blue
Connect-SPOService -Url $SharePointAdminURL -Credential $credentials
  
# Set current admin as a Site Collection Admin on both OneDrive Site Collections
Write-Host "`nAdding $globaladmin as site collection admin on both OneDrive site collections" -ForegroundColor Blue
Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
Set-SPOUser -Site $DestinationSharepointSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
  
Write-Host "`nConnecting to SharePoint Online PNP module" -ForegroundColor Blue

$ConnectionSite = Connect-PnPOnline -url $DestinationSharepointSite -Credentials $credentials -ReturnConnection
$ConnectionDrive = Connect-PnPOnline -url $departingOneDriveSite -Credentials $Credentials

# Get name of departing user to create folder name.
Write-Host "`nGetting display name of $departinguser" -ForegroundColor Blue
$departingOwner = Get-PnPSiteCollectionAdmin | Where-Object {$_.loginname -match $departinguser}
  
# If there's an issue retrieving the departing user's display name, set this one.
if ($departingOwner -contains $null) {
    $departingOwner = @{
        Title = "Departing User"
    }
}

# Define relative folder locations for OneDrive source and destination
$departingOneDrivePath = "/personal/$departingUserUnderscore/Documents"
$destinationSitePath = "/sites/$DestinationSiteFormatted/shared Documents/$($departingOwner.Title)'s Files"
$DestinationSharepointSiteRelativePath = "shared Documents/$($departingOwner.Title)'s Files"

# Get all items from source OneDrive
Write-Host "`nGetting all items from $($departingOwner.Title)" -ForegroundColor Blue
$items = Get-PnPListItem -List Documents -PageSize 1000 -Connection $ConnectionDrive



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

# Filter by Folders to create directory structure
Write-Host "`nFilter by folders" -ForegroundColor Blue
$folders = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "Folder"}
$i = 0
$total = $folders.Count
Write-Host "`nCreating Directory Structure" -ForegroundColor Blue
foreach ($folder in $folders) {
    $path = ('{0}{1}' -f $DestinationSharepointSiteRelativePath, $folder.fieldvalues.FileRef) -Replace $departingOneDrivePath
    $i++
    Write-Progress -Activity "Creating Directory Structure" -PercentComplete (($i/$total)*100) -CurrentOperation "Creating folder in $path"
    #Write-Host "Creating folder in $path" -ForegroundColor Green
    $newfolder = Resolve-PnPFolder -SiteRelativePath $path -Connection $ConnectionSite
}

Write-Host "`nCopying Files" -ForegroundColor Blue
$LocalPath = "c:\temp"  
$files = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "File"}
$fileerrors = ""
$i = 0
$total = $files.Count
foreach ($file in $files) {
    $name = $file.fieldvalues.FileLeafRef.tostring()   
    $destpath = ("$destinationSitePath$($file.fieldvalues.FileDirRef)") -Replace $departingOneDrivePath
    $path = $file.fieldvalues.FileRef
    $i++
    Write-Progress -Activity "Creating Directory Structure" -PercentComplete (($i/$total)*100) -CurrentOperation "Copying $($file.fieldvalues.FileLeafRef) to $destpath"
    #Write-Host "Copying $($file.fieldvalues.FileLeafRef) to $destpath" -ForegroundColor Green
    Get-PnPFile -Url $path -Path $LocalPath -Filename $name -AsFile -force -Connection $ConnectionDrive
    $newfile = Add-PnPFile -Path "$LocalPath\$name" -folder $destpath -Connection $Connectionsite
    Remove-Item -Path "$localPath\$name" -Force
}
$fileerrors | Out-File c:\temp\fileerrors.txt
  
# Remove Global Admin from Site Collection Admin role for both users
Write-Host "`nRemoving $globaladmin from OneDrive site collections" -ForegroundColor Blue
Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
Set-SPOUser -Site $DestinationSharepointSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
Write-Host "`nComplete!" -ForegroundColor Green
