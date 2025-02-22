<#
.SYNOPSIS
    Send Message to Teams users.
.DESCRIPTION
    Uses the Graph API to send messages to users (all or specific) of an organization. The message is based on a template.
.NOTES
    To use this, just download all files from the top left corner as zip, extract the content. Change the message template and run one of the examples.
.EXAMPLE
PS C:\> .\Send-TeamsMessage.ps1
cmdlet Send-TeamsMessage.ps1 at command pipeline position 1
Supply values for the following parameters:
MessagePath: Send-TeamsMessage.txt
UserEmail[0]: name@domain.com
UserEmail[1]:
Message sent to name@domain.com - 1712341085558
.EXAMPLE
PS C:\> .\Send-TeamsMessage.ps1 -All
cmdlet Send-TeamsMessage.ps1 at command pipeline position 1
Supply values for the following parameters:
MessagePath: Send-TeamsMessage.txt
Message sent to name1@domain.com - 171234108555
Message sent to name2@domain.com - 171234101238
Message sent to name3@domain.com - 171231234558
Message sent to name4@domain.com - 121443105558
#>

[CmdletBinding(DefaultParameterSetName = "Specific")]
param (
    # The path to the file containing the message to send. {{GivenName}} will be replaced with the user's first name. Use [Linktext](https://link.de) to create links.
    [Parameter(Mandatory)]
    [System.IO.FileInfo]
    $MessagePath,

    # The Mail address of the user to send message to.
    [Parameter(ParameterSetName = "Specific", Mandatory)]
    [String[]]
    $UserEmail,

    # Send message to all users.
    [Parameter(ParameterSetName = "All", Mandatory)]
    [switch]$All
)

#requires -Version 7
#requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Teams

#region Global Variables
$ErrorActionPreference = "Stop"
#endregion

#region Connect
Connect-MgGraph -Scopes "User.ReadBasic.All", "Chat.Create", "ChatMessage.Send"
#endregion

#region Get users
$users = Get-MgUser -All
$context = Get-MgContext

if ($PSCmdlet.ParameterSetName -eq "Specific") {
    $users = $users | Where-Object { $UserEmail -contains $_.Mail }
}
#endregion

#region send message

$message = Get-Content -LiteralPath $MessagePath -Raw

foreach ($user in $users) {
    $params = @{
        chatType = "oneOnOne"
        members  = @(
            @{
                "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                roles             = @(
                    "owner"
                )
                "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($user.Id)')"
            }
            @{
                "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                roles             = @(
                    "owner"
                )
                "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($context.Account)')"
            }
        )
    }
    
    $chat = New-MgChat -BodyParameter $params

    $userMessage = $Message -replace "{{GivenName}}", $user.GivenName
    $body = @{
        body = @{
            content     = $userMessage
            contentType = "html"
        }
    }
    $chatMessage = New-MgChatMessage -ChatId $chat.Id -BodyParameter $body

    Write-Information -InformationAction Continue -MessageData "Message sent to $($user.Mail) - $($chatMessage.Id)"
}
