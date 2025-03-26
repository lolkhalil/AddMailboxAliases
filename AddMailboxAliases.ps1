# AddMailboxAliases - Script by Khalil
# Used to add Multiple Aliases to an Email Account

function WriteR($message) { Write-Host -ForegroundColor "Red" $message }
function WriteG($message) { Write-Host -ForegroundColor "Green" $message }
function WriteY($message) { Write-Host -ForegroundColor "Yellow" $message }

# Module Check to ensure Exchange is installed
$Modules = @("ExchangePowerShell","ExchangeOnlineManagement")
try {
    foreach ($Module in $Modules) {
        if ($Module -notin (Get-InstalledModule).Name) {
            Write-Host "Installing $Module"
            Install-Module -Name $Module -Force | Out-Null
        }
    }
} catch [Exception] {
    Write-Host "Couldn't install Powershell Modules"
    Read-Host
    $_ # This will rarely happen so an error output would be useful
    exit
}

# Connecting to Exchange
try {
    Connect-ExchangeOnline -ErrorAction Stop | Out-Null
} catch {
    WriteR "Couldn't connect to Exchange"
    Read-Host
    exit
}

# continue with rest of the stuff

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