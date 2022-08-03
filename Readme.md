This script will move all folders and files from a users onedrive to a sharepoint site

The script will ask for three inputs:

  * The Users email
  * The site name 
      * ex. domain.sharepoint/sites/legal --> legal
  * An Admin account email
      * A credentials box will pop up asking for the password

To install the dependencies for this script use these two powershell commands in an elevated terminal

`Install-Module -Name MSOnline`

`Install-Module -Name "PnP.PowerShell" `
