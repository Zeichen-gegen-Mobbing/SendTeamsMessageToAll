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

When you want to send mail to all people type `-All` afterwards. You can exclude specific people with `-ExcludeDisplayName "Karla Kolumna"`.

### Prerequistes

You need PowerShell 7. To install it run `winget install Microsoft.PowerShell`.

Afterwards open Powershell (with the black logo) and run `Install-Module Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Teams -Scope CurrentUser` to install required modules
