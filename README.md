# ModifyMailboxAliases
Adding/Removing Aliases to an Email Account from a List on a CSV File

## Usage:

```Powershell
# Add Aliases
.\ModifyMailboxAliases -Email email@domain.com -CSVFilePath Aliases.csv -Option Add

# Remove Aliases
.\ModifyMailboxAliases -Email email@domain.com -CSVFilePath Aliases.csv -Option Remove
```