#create contacts based on export from Tomra
# run from exchange online ps module

$csvfile = '.\tomrausers.csv'

if(!(Get-AcceptedDomain).domainname -contains 'compacsort.com'){
    Connect-EXOPSSession # use credentials from target tenant (Compac)
}

$SourceUsers = Import-Csv -Path $csvfile -Delimiter ";" -Encoding Default | ?{$_.windowsemailaddress -like "*@tomra.com"}
$createcount = 0

$existingcontacts = Get-MailContact -ResultSize unlimited
$existingguests = Get-MailUser -ResultSize unlimited | ?{$_.primarysmtpaddress -ne $null} #some users added as guests, catch them here to avoid error creating contacts

$ContactsTable =@{}
$existingcontacts | %{
    $ContactsTable.Add($_.windowsemailaddress,$_)
}

$existingguests | %{
    $ContactsTable.Add($_.primarysmtpaddress,$_)
}

foreach($user in $SourceUsers){
    #Write-Host sjekker $($user.WindowsEmailAddress)
    if(!$ContactsTable.Contains($($user.WindowsEmailAddress))){
        #doesn't exist - Create
        Write-Host Oppretter $($user.WindowsEmailAddress)
        New-MailContact -Name $($user.DisplayName) -LastName $($user.LastName) -DisplayName $($user.DisplayName) -ExternalEmailAddress $($user.WindowsEmailAddress) -FirstName $($user.FirstName)
        $sip = "sip:" + $user.WindowsEmailAddress
        Set-MailContact $($user.WindowsEmailAddress) -EmailAddresses @{add=$sip}
        $createcount += 1
    }
}

Write-Host "Created $createcount new contacts"
