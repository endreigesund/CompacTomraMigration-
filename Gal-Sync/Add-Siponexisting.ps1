$existingcontacts = Get-MailContact -ResultSize unlimited -Filter {windowsemailaddress -like "*tomra.com"}

$existingcontacts | %{
    $sip = "sip:" + $_.windowsemailaddress
    if($_.emailaddresses -notcontains $sip){
        Write-Host $sip - legger til
        $_ | Set-MailContact -EmailAddresses @{add=$sip}
    }
    else{Write-Host $sip - ok}
}