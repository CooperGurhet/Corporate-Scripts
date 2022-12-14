function Send-HTMLEmail {
    #Requires -Version 3
    [CmdletBinding()]
     Param 
       ([Parameter(Mandatory=$True,
                   Position = 0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please enter the Inputobject")]
        $InputObject,
        [Parameter(Mandatory=$True,
                   Position = 1,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please enter the Subject")]
        [String]$Subject,    
        [Parameter(Mandatory=$False,
                   Position = 2,
                   HelpMessage="Please enter the To address")]    
        [String[]]$To = '222c70a6.prometric.onmicrosoft.com@amer.teams.ms',
        [String]$From = 'Office365.Admin@prometric.com',    
        [String]$CSS,
        [String]$SmtpServer ="promailervip.proint.prometric.root"
       )#End Param
    
    if (!$CSS)
    {
        $CSS = @"
            <style type="text/css">
                table {
                font-family: Verdana;
                border-style: dashed;
                border-width: 1px;
                border-color: #FF6600;
                padding: 5px;
                background-color: #FFFFCC;
                table-layout: auto;
                text-align: center;
                font-size: 8pt;
                }
    
                table th {
                border-bottom-style: solid;
                border-bottom-width: 1px;
                font: bold
                }
                table td {
                border-top-style: solid;
                border-top-width: 1px;
                }
                .style1 {
                font-family: Courier New, Courier, monospace;
                font-weight:bold;
                font-size:small;
                }
                </style>
"@
    }#End if
    
    $HTMLDetails = @{
        Title = $Subject
        Head = $CSS
        }
        
    $Splat = @{
        To         =$To
        Body       ="$($InputObject | ConvertTo-Html @HTMLDetails)"
        Subject    =$Subject
        SmtpServer =$SmtpServer
        From       =$From
        BodyAsHtml =$True
        }
        Send-MailMessage @Splat -Priority High
        
}
$starttime = Get-Date
$departinguser = Read-Host "Enter departing user's email"
$DestinationSite = Read-Host "Enter destination site Name"
$destinationsubsite = Read-host "Enter Destination sub site name"
$ScriptEmail = Read-Host "Enter your Email"
$Requestinguser = Read-Host "Enter Requesting user's email"
$globaladmin = Read-Host "Enter the username of your Global Admin account"
$credentials = Get-Credential -Credential $globaladmin
Connect-MsolService -Credential $credentials
 
$InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}
  
$SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"
  
$departingUserUnderscore = $departinguser -replace "[^a-zA-Z]", "_"
$DestinationSiteFormatted = $DestinationSite -replace "[^a-zA-Z]", ""
$DestinationsubSiteFormatted = $DestinationsubSite -replace "[^a-zA-Z]", ""
  
$departingOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$departingUserUnderscore"
$DestinationSharepointSite = "https://$($InitialDomain.Name.Split(".")[0]).sharepoint.com/sites/$DestinationSiteFormatted/$Destinationsubsiteformatted"

Write-Host "`nConnecting to SharePoint Online" -ForegroundColor Green
Connect-SPOService -Url $SharePointAdminURL -Credential $credentials
  
# Set current admin as a Site Collection Admin on both OneDrive Site Collections
Write-Host "`nAdding $globaladmin as site collection admin on both OneDrive site collections" -ForegroundColor Green
Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
Set-SPOUser -Site $DestinationSharepointSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
  
Write-Host "`nConnecting to SharePoint Online PNP module" -ForegroundColor Green

$ConnectionSite = Connect-PnPOnline -url $DestinationSharepointSite -Credentials $credentials -ReturnConnection
$ConnectionDrive = Connect-PnPOnline -url $departingOneDriveSite -Credentials $Credentials

# Get name of departing user to create folder name.
Write-Host "`nGetting display name of $departinguser" -ForegroundColor Green
$departingOwner = Get-PnPSiteCollectionAdmin | Where-Object {$_.loginname -match $departinguser}
  
# If there's an issue retrieving the departing user's display name, set this one.
if ($departingOwner -contains $null) {
    $departingOwner = @{
        Title = "Departing User"
    }
}

# Define relative folder locations for OneDrive source and destination
$departingOneDrivePath = "/personal/$departingUserUnderscore/Documents"
$destinationSitePath = "/sites/$DestinationSiteFormatted/$DestinationsubsiteFormatted/shared Documents/$($departingOwner.Title)'s Files"
$DestinationSharepointSiteRelativePath = "shared Documents/$($departingOwner.Title)'s Files"

# Get all items from source OneDrive
Write-Host "`nGetting all items from $($departingOwner.Title)" -ForegroundColor Green
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
Write-Host "`nFilter by folders" -ForegroundColor Green
$folders = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "Folder"}
$i = 0
$total = $folders.Count
Write-Host "`nCreating Directory Structure" -ForegroundColor Green
foreach ($folder in $folders) {
    $path = ('{0}{1}' -f $DestinationSharepointSiteRelativePath, $folder.fieldvalues.FileRef) -Replace $departingOneDrivePath
    $i++
    Write-Progress -Activity "Creating Directory Structure" -status "$i/$total" -PercentComplete (($i/$total)*100) -CurrentOperation "Creating folder in $path"
    #Write-Host "Creating folder in $path" -ForegroundColor Green
    $newfolder = Resolve-PnPFolder -SiteRelativePath $path -Connection $ConnectionSite
}

Write-Host "`nCopying Files" -ForegroundColor Green
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
    Write-Progress -Activity "Copying Files" -status "$i/$total" -PercentComplete (($i/$total)*100) -CurrentOperation "Copying $($file.fieldvalues.FileLeafRef) to $destpath"
    #Write-Host "Copying $($file.fieldvalues.FileLeafRef) to $destpath" -ForegroundColor Green
    Get-PnPFile -Url $path -Path $LocalPath -Filename $name -AsFile -force -Connection $ConnectionDrive
    $newfile = Add-PnPFile -Path "$LocalPath\$name" -folder $destpath -Connection $Connectionsite
    Remove-Item -Path "$localPath\$name" -Force
}
$fileerrors | Out-File c:\temp\fileerrors.txt
  
# Remove Global Admin from Site Collection Admin role for both users
Write-Host "`nRemoving $globaladmin from OneDrive site collections" -ForegroundColor Green
Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
Set-SPOUser -Site $DestinationSharepointSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
Write-Host "`nComplete!" -ForegroundColor Green
$endtime = Get-Date
Write-Host "started at: $starttime"
write-host "completed at: $endtime"
Send-HTMLEmail -To "$($ScriptEmail); $($Requestinguser)" -Subject "$($departinguser)'s One Drive Files have been moved to $($DestinationSharepointSite)" -Inputobject "The files have been moved to https://$($InitialDomain.Name.Split(".")[0]).sharepoint.com/shared Documents/$($departingOwner.Title)'s Files"