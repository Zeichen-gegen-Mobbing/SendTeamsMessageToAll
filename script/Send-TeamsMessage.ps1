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
UserEmail[0]: user1@example.com
UserEmail[1]:
Message sent to user1@example.com - 1712341085558
.EXAMPLE
PS C:\> .\Send-TeamsMessage.ps1 -UserEmail "user1@example.com","user2@example.com"
Message sent to user1@example.com - 1712341085558
Message sent to user2@example.com - 1712321055558
.EXAMPLE
PS C:\> .\Send-TeamsMessage.ps1 -Group "My Team"
Message sent to user1@example.com - 1712341085558
Message sent to user2@example.com - 1712321055558
.EXAMPLE
PS C:\> .\Send-TeamsMessage.ps1 -All
Message sent to user1@example.com - 1712341085558
Message sent to user2@example.com - 1712321055558
Message sent to user3@example.com - 121443105558
.EXAMPLE
PS C:\> .\Send-TeamsMessage.ps1 -All -ExcludeDisplayName "Ernie Sesame","Karla Kolumna"
Message sent to user1@example.com - 1712341085558
Message sent to user2@example.com - 1712321055558
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
    [Parameter(ParameterSetName = "Specific", Mandatory, Position = 0)]
    [String[]]
    $UserEmail,

    # Send message to all members (no guests).
    [Parameter(ParameterSetName = "All", Mandatory)]
    [switch]$All,

    # Send message to members of a specific Team. When you use this, you must have a Admin role.
    [Parameter(ParameterSetName = "Group", Mandatory)]
    [ValidateNotNullOrEmpty()]
    [Alias("Team")]
    [string]$Group,


    # Exclude members by DisplayName. When you use this, you must have a Admin role. 
    [Parameter(ParameterSetName = "All")]
    [Parameter(ParameterSetName = "Group")]
    [String[]]
    $ExcludeDisplayName
)

#requires -Version 7
#requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Groups,Microsoft.Graph.Teams

#region Global Variables
$ErrorActionPreference = "Stop"
#endregion

#region Connect

$additionalScope = "User.ReadBasic.All"
if ($PSCmdlet.ParameterSetName -eq "Group") {
    $additionalScope = "GroupMember.Read.All"
    Write-Information -InformationAction Continue -MessageData "You are using the Group parameter. This requires the Group.Read.All scope, which is an Admin scope. Make sure you have the necessary permissions to run this script."
    Write-Information -InformationAction Continue -MessageData "If you don't have the Group.Read.All scope, you can remove the Group parameter to use the User.ReadBasic.All scope instead."
}
elseif ($PSBoundParameters.ContainsKey("ExcludeDisplayName")) {
    $additionalScope = "User.Read.All"
    Write-Information -InformationAction Continue -MessageData "You are using the ExcludeDisplayName parameter. This requires the User.Read.All scope, which is an Admin scope. Make sure you have the necessary permissions to run this script."
    Write-Information -InformationAction Continue -MessageData "If you don't have the User.Read.All scope, you can remove the ExcludeDisplayName parameter to use the User.ReadBasic.All scope instead."
}

$scopes = @("Chat.Create", "ChatMessage.Send", $additionalScope)
Connect-MgGraph -Scopes $scopes
#endregion

#region Get users
$properties = @("id", "givenName", "mail", "userType")
if ($PSBoundParameters.ContainsKey("ExcludeDisplayName")) {
    $properties += "displayName"
}

if ($PSCmdlet.ParameterSetName -eq "Group") {
    $groupId = (Get-MgGroup -Filter "displayName eq '$Group'" -Property "id").Id
    if (-not $groupId) {
        throw "Group '$Group' not found. Please check the group name."
    }
    $users = Get-MgGroupMember -GroupId $groupId -All -Property $properties
}
else {
    $users = Get-MgUser -All -Property $properties
    $context = Get-MgContext

    if ($PSCmdlet.ParameterSetName -eq "Specific") {
        $users = $users | Where-Object { $UserEmail -contains $_.Mail }
    }
    elseif ($PSCmdlet.ParameterSetName -eq "All") {
        $users = $users | Where-Object { $_.UserType -ne "Guest" }
    }
}

if ($PSBoundParameters.ContainsKey("ExcludeDisplayName")) {
    $users = $users | Where-Object { $ExcludeDisplayName -notcontains $_.DisplayName }
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
    # $chatMessage = New-MgChatMessage -ChatId $chat.Id -BodyParameter $body

    Write-Information -InformationAction Continue -MessageData "Message sent to $($user.Mail) - $($chatMessage.Id)"
}
