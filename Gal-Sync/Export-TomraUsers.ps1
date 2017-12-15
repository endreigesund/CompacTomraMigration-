#eksport all user mailbox to a csv-file for contact creation in another tenant
# run from exchange online ps module

$csvfile = '.\tomrausers.csv'

if(!(Get-AcceptedDomain).domainname -contains 'tomra.com'){
    Connect-EXOPSSession # use credentials from source tenant (tomra)
}
    
$allmbx = get-user -RecipientTypeDetails usermailbox -ResultSize unlimited -Filter {customattribute12 -eq $null}

$allmbx | Select-Object City, Company, Department, DisplayName, FirstName, LastName, MobilePhone, Office, Phone, Title, WindowsEmailAddress | Export-Csv -Path $csvfile -Encoding Default -NoClobber -NoTypeInformation -Delimiter ";"

