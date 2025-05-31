# SendTeamsMessageToAll

Send Teams Message to multiple people using pwsh.

The files are also available in this repository in the folder "scripts".
There are two important files:

- Send-TeamsMessage.ps1 - PowerShell Script
- MessageTemplate.txt - Template for messages in HTML

## Usage

First change the [Template](./script/MessageTemplate.txt) to your liking.

Secondly right click on the `script/Send-Teams-Message` file and select copy as path.

Lastly open Powershell (with the black logo). Type a `.` and paste the copied path afterwards.

If you want to send Messages to specific people press enter afterwards and enter mail address after mail address.
You can also add multiple mail addresses (like from excel) `"user1@example.com","user2@example.com"`.

When you want to send mail to all people type `-All` afterwards. You can exclude specific people with `-ExcludeDisplayName "Karla Kolumna"`.

### Prerequistes

You need PowerShell 7. To install it run `winget install Microsoft.PowerShell`.

Afterwards open Powershell (with the black logo) and run `Install-Module Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Teams,Microsoft.Graph.Groups -Scope CurrentUser` to install required modules

### Troubleshooting

On some devices, the following error message occurs: " Die Datei "..." kann nicht geladen werden, da die Ausführung von Skripts auf diesem System deaktiviert ist.". To solve this problem, you have to exclude the script from your execution policy.
To do so, run `Unblock-File -Path "C:\path\to\the\script"` in your PowerShell window. You have to replace the default path with the one you copied as a path above when trying to run the script.
You might also try to bypass the execution policy if it is not `RemoteSigned`. Run `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process` to disable the execution policy for the current process
