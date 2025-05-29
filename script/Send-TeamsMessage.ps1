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
.EXAMPLE
PS C:\> .\Send-TeamsMessage.ps1 -All -ExcludeDisplayName "Ernie Sesame","Karla Kolumna"
cmdlet Send-TeamsMessage.ps1 at command pipeline position 1
Supply values for the following parameters:
MessagePath: Send-TeamsMessage.txt
Message sent to name1@domain.com - 171234108555
Message sent to name4@domain.com - 121443105558
#>

[CmdletBinding(DefaultParameterSetName = "Specific")]
param (
    # The path to the file containing the message to send. {{GivenName}} will be replaced with the user's first name. Use [Linktext](https://link.de) to create links.
    [System.IO.FileInfo]
    [ValidateScript({
            Test-Path $_ -PathType Leaf
        })]
    $MessagePath = (Join-Path -Path $PSScriptRoot -ChildPath "MessageTemplate.txt"),

    # The Mail address of the user in current Tenant to send message to.
    [Parameter(ParameterSetName = "Specific", Mandatory)]
    [String[]]
    $UserEmail,

    # Send message to all members (no guests).
    [Parameter(ParameterSetName = "All", Mandatory)]
    [switch]$All,

    # Exclude members by DisplayName. When you use this, you must have a Admin role. 
    [Parameter(ParameterSetName = "All")]
    [String[]]
    $ExcludeDisplayName
)

#requires -Version 7
#requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Teams

#region Global Variables
$ErrorActionPreference = "Stop"
#endregion

#region Connect
$scopes = @("Chat.Create", "ChatMessage.Send", "")
if ($PSBoundParameters.ContainsKey("ExcludeDisplayName")) {
    $scopes[2] = "User.Read.All"
    Write-Warning -MessageData "You are using the ExcludeDisplayName parameter. This requires the User.Read.All scope, which is an Admin scope. Make sure you have the necessary permissions to run this script."
    Write-Information -InformationAction Continue -MessageData "If you don't have the User.Read.All scope, you can remove the ExcludeDisplayName parameter to use the User.ReadBasic.All scope instead."
}
else {
    $scopes[2] = "User.ReadBasic.All"
}

Connect-MgGraph -Scopes $scopes
#endregion

#region Get users
$users = Get-MgUser -All -Property id, givenName, mail, userType, displayName
$context = Get-MgContext

if ($PSCmdlet.ParameterSetName -eq "Specific") {
    $users = $users | Where-Object { $UserEmail -contains $_.Mail }
}
elseif ($PSCmdlet.ParameterSetName -eq "All") {
    $users = $users | Where-Object { $_.UserType -ne "Guest" -and $_.DisplayName -notin $ExcludeDisplayName }
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
