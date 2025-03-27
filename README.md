# AddMailboxAliases
Adding a number of Aliases to an Email Account

## Usage:

```Powershell
# Add Aliases
.\AddMailboxAliases -Email MAILBOX@DOMAIN -CSVFilePath FILENAME.csv -Option Add

# Remove Aliases
.\AddMailboxAliases -Email MAILBOX@DOMAIN -CSVFilePath FILENAME.csv -Option Remove
```