# Adding 100 Aliases to an account

$array = @()
foreach ($num in 1..100) {
    $array += "TestMailbox" + $num + "@REDACTED.onmicrosoft.com"
}

$Mailbox = "TestMailbox@REDACTED.onmicrosoft.com"
foreach ($mb in $array) { 
    Set-Mailbox $Mailbox -EmailAddresses @{add=$mb} 
}

foreach ($mb in $array) { 
    Set-Mailbox $Mailbox -EmailAddresses @{remove=$mb} 
}