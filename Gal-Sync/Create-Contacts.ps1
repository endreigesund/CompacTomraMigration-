#create contacts based on export from Tomra
# run from exchange online ps module

$csvfile = 'C:\temp\tomrausers.csv'

Connect-EXOPSSession # use credentials from target tenant (Compac)

$SourceUsers = Import-Csv -Path $csvfile -Delimiter ";" -Encoding Default | ?{$_.windowsemailaddress -like "*@tomra.com"}
$createcount = 0

foreach($user in $SourceUsers){
    if((Get-Recipient $($user.WindowsEmailAddress) -ErrorAction silentlycontinue) -eq $null){
        #doesn't exist - Create
        New-MailContact -Name $($user.DisplayName) -LastName $($user.LastName) -DisplayName $($user.DisplayName) -ExternalEmailAddress $($user.WindowsEmailAddress) -FirstName $($user.FirstName)
        $createcount += 1
    }
}

Write-Host "Created $createcount new contacts"
