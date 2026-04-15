# Import PowerShell modules for Microsoft Graph and Exchange Online
Import-Module Microsoft.Graph.Authentication -SkipEditionCheck
Import-Module Microsoft.Graph.Users -SkipEditionCheck
Import-Module Microsoft.Graph.Groups -SkipEditionCheck
Import-Module Microsoft.Graph.Applications -SkipEditionCheck
Import-Module ExchangeOnlineManagement -SkipEditionCheck


$GroupId = 'GROUP-ID-HERE'
$DomainName = 'DOMAIN-NAME-HERE'

# Retrieve authentication token from Graph, convert to secure string.
Connect-AzAccount -Identity | Out-Null
$Token = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
$SecureToken = ($Token.Token | ConvertTo-SecureString -AsPlainText -Force) # Convert plain text token to secure string (required in Graph API v2)

# Connect to Microsoft Graph session using least privilege.
Connect-MgGraph -AccessToken $SecureToken -NoWelcome

# Connect to Exchange Online using managed identity (permissions assigned in Entra ID with Graph API)
Connect-ExchangeOnline -ManagedIdentity -Organization $DomainNAme

# Retrieve list of shared mailboxes and export to csv.
$SharedMailboxes = Get-EXORecipient -ResultSize unlimited -RecipientTypeDetails "SharedMailbox" 

# Retrieve group members of security group.
$GroupMembership = Get-MgGroupMember -GroupId $GroupId -Top 10000 | ForEach-Object {
    [pscustomobject]@{
        Id                 = $_.id
        displayName        = $_.additionalproperties['displayName']
        PrimarySMTPAddress = $_.additionalproperties['mail']
        userPrincipalName  = $_.additionalproperties['userPrincipalName']
    }
}

# Compares list of existing shared mailboxes found in exchange against Entra ID security group membership.
$AddMembers = $SharedMailboxes | Where-Object { $_.PrimarySMTPAddress -NotIn $GroupMembership.PrimarySMTPAddress }

if ($AddMembers)
{
    # Retrieves object ID of each user account requiring membership to Entra ID security group.
    $AddUserObjectId = ($AddMembers | ForEach-Object { Get-MgUser -Filter "Mail eq '$($_.PrimarySmtpAddress)'" }).Id
    # Adds each user as member to the Entra ID security group containing shared mailboxes.
    $AddUserObjectId | ForEach-Object { New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $_ }
}


# Compares list of existing group members found in Entra ID security group and compares against existing shared mailboxes found in Exchange.
$RemoveMembers = $GroupMembership | Where-Object { $_.PrimarySMTPAddress -NotIn $SharedMailboxes.PrimarySMTPAddress }

if ($RemoveMembers)
{
    # Retrieves object ID of each user account requiring membership removal to Entra ID security group.
    $RemoveUserObjectId = ($RemoveMembers | ForEach-Object { Get-MgUser -Filter "Mail eq '$($_.PrimarySmtpAddress)'" }).Id

    # Removes each user as member to the Entra ID security group containing shared mailboxes.
    $RemoveUserObjectId | ForEach-Object { Remove-MgGroupMemberDirectoryObjectByRef -GroupId $GroupId -DirectoryObjectId $_ }
}

# Disconnect Exchange Online session.
Disconnect-ExchangeOnline -Confirm:$false

# Disconnect Microsoft Graph session.
Disconnect-MgGraph
