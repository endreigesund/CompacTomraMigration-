#create contacts based on export from Tomra
# run from exchange online ps module

$csvfile = 'C:\temp\tomrausers.csv'

Connect-EXOPSSession # use credentials from target tenant (Compac)

$SourceUsers = Import-Csv -Path $csvfile -Delimiter ";" -Encoding Default | ?{$_.windowsemailaddress -like "*@tomra.com"}
$createcount = 0

$existingcontacts = Get-MailContact -ResultSize unlimited

$ContactsTable =@{}
$existingcontacts | %{
    $ContactsTable.Add($_.windowsemailaddress,$_)

}

foreach($user in $SourceUsers){
    Write-Host sjekker $($user.WindowsEmailAddress)
    if(!$ContactsTable.Contains($($user.WindowsEmailAddress))){
        #doesn't exist - Create
        Write-Host Oppretter $($user.WindowsEmailAddress)
        New-MailContact -Name $($user.DisplayName) -LastName $($user.LastName) -DisplayName $($user.DisplayName) -ExternalEmailAddress $($user.WindowsEmailAddress) -FirstName $($user.FirstName)
        $createcount += 1
    }
}

Write-Host "Created $createcount new contacts"
