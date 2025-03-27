# ModifyMailboxAliases - Script by Khalil
# Used to add/remove Multiple Aliases to an Email Account

# Usage: .\ModifyMailboxAliases -Email MAILBOX@DOMAIN -CSVFilePath FILENAME.csv -Option Add
#        .\ModifyMailboxAliases -Email MAILBOX@DOMAIN -CSVFilePath FILENAME.csv -Option Remove
param($Email, $CSVFilePath, $Option)

function WriteR($message) { Write-Host -ForegroundColor "Red" $message }
function WriteG($message) { Write-Host -ForegroundColor "Green" $message }
function WriteY($message) { Write-Host -ForegroundColor "Yellow" $message }

$LogFile = "Aliases.log"
function AppendLogFile($Message) {
    if (!(Test-Path -Path $LogFile)) {
        New-Item -Path $LogFile | Out-Null
    }

    Add-Content $LogFile -Value $Message
}

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
    $_ # This will rarely happen so an error output would be useful
    exit
}

# Connecting to Exchange
try {
    Connect-ExchangeOnline -ErrorAction Stop *> $null
} catch {
    WriteR "Couldn't connect to Exchange Online"
    exit
}

# Getting the correct Option (Add/Remove)
while ($null -eq $Option) {
    $Option = Read-Host "Would you like to Add or Remove Aliases?"
    
    if ($Option -notlike "add" -and $Option -notlike "remove") {
        WriteR "Provide a correct input. Add/Remove"
        $Option = $null
    }
}

# Getting the Path of the CSV File of Aliases
while ($null -eq $CSVFilePath) {
    $CSVFilePath = Read-Host "Where is the CSV File of Aliases?"
    
    try {
        if (Test-Path -Path $CSVFilePath) {
            break;
        }
    } catch {}
    
    WriteR "Couldn't find CSV File at the Path: $CSVFilePath"
    $CSVFilePath = $null
}

# Used to Import the CSV File
try {
    $CSVFile = Import-Csv $CSVFilePath
    if ($null -eq $CSVFile) {
        WriteR "Couldn't find any data in the CSV File: $CSVFilePath"
        Disconnect-ExchangeOnline -Confirm:$false
        exit
    }
} catch {
    WriteR "Couldn't Import the CSV File at the Path: $CSVFilePath"
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Getting the Email Address of the Mailbox
while ($true) {
    if ($null -eq $Email) {
        $Email = Read-Host "What is the Mailbox's Email Address?"
    }
    
    # Checking if the Mailbox exists
    try {
        Get-Mailbox -Identity $Email -ErrorAction Stop *> $null
        break
    } catch {
        WriteR "Couldn't find the Mailbox: $Email"
        $Email = $null
    }
}

AppendLogFile("============================")
AppendLogFile("Script Started: $((Get-Date).ToString())")

# The Name property is the Header, so it gets all the information from the Header they used
$Aliases = $CSVFile.($CSVFile[0].PSObject.Properties.Name)
# Starting to Add/Remove the Aliases
foreach ($Alias in $Aliases) { 
    try {
        Set-Mailbox $Email -EmailAddresses @{$Option=$Alias}
        AppendLogFile("Alias: $($Alias): $($Option) to/from Mailbox: $($Email)")
        WriteG "$Option Alias: $Alias to/from Mailbox: $Email"
    } catch [Exception] {
        AppendLogFile("ERROR: Alias: $($Alias): Couldn't $($Option) to/from Mailbox: $($Email) ($_)")
        WriteR("Couldn't $Option Alias: $Alias to/from Mailbox: $Email")
    }
}

try {
    $ProxyAddresses = (Get-Mailbox $Email).EmailAddresses
} catch {
    WriteR "Couldn't get the Mailbox Proxy Addresses: $Email"
    exit
} finally {
    Disconnect-ExchangeOnline -Confirm:$false
}

Write-Host ""
WriteY "Current Proxy Addresses: "
foreach ($Address in $ProxyAddresses) { WriteY $Address }

# foreach ($num in 1..100) {
#     $array += "TestMailbox" + $num + "@REDACTED.onmicrosoft.com"
# }