Connect-ExchangeOnline

$Name = read-host "Enter Distribution List Name"
$Managedby = read-host "Enter email of Distribution List Manager" #Comment out if Distribution List is already Created
$members = ((get-content .\test.csv).repace(" ","")).Split(',')  #Create a file called test.csv in the same directory as script, Seperate each emails with a comma

New-DistributionGroup -Name $Name -ManagedBy $Managedby -PrimarySmtpAddress "$($Name)@prometric.com" -type distribution  #Comment out if Distribution List is already Created
Set-DistributionGroup -Identity $Name -AcceptMessagesOnlyFrom $Managedby #Comment out if Distribution List is already Created
foreach($member in $members){
    Add-DistributionGroupMember -Identity $Name -Member $member
}
