#Update existing contacts based on export from Tomra
# run from exchange online ps module

$csvfile = 'C:\temp\tomrausers.csv'

if(!(Get-AcceptedDomain).domainname -contains 'compacsort.com'){
    Connect-EXOPSSession # use credentials from target tenant (Compac)
}

$SourceUsers = Import-Csv -Path $csvfile -Delimiter ";" -Encoding Default | ?{$_.windowsemailaddress -like "*@tomra.com"}
$updatecount = 0

$TargetContacts = Get-Contact -ResultSize unlimited | ?{$_.windowsemailaddress -like "*@tomra.com"}

#hash tomrausers for performance
$SourceUsersTable =@{}
$SourceUsers | %{
    $SourceUsersTable.Add($_.windowsemailaddress,$_)

}

#define properties to compare
$props = $SourceUsers | Get-Member | ?{$_.membertype -like 'Noteproperty'}

foreach($Contact in $TargetContacts){
    $update = $false
    $SourceUser = $SourceUsersTable.Item($Contact.windowsemailaddress)
    if($SourceUser){
        Write-Host Sammenligner info på bruker - $($Contact.windowsemailaddress)
        foreach($property in $props){
            if($SourceUser.$($property.name) -notlike $Contact.$($property.name)){
                Write-Host Forskjell i $($property.name) - må oppdatere.
                $update = $true
            }
        }
        if($update){
            $contact | Set-Contact -City $($sourceuser.city) -Company $($sourceuser.Company) -Department $($sourceuser.Department) -DisplayName $($sourceuser.DisplayName) -FirstName $($sourceuser.FirstName) -LastName $($sourceuser.LastName) -MobilePhone $($sourceuser.MobilePhone) -Office $($sourceuser.Office) -Phone $($sourceuser.Phone) -Title $($sourceuser.Title)
            $updatecount += 1
        }
    }

}

Write-Host "Updated $updatecount contacts"