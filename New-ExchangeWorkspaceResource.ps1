<#

.EXAMPLE

This script was desgined to be launched using a POST API request within the context of a workflow in Freshservice Workflow Automator, however, it can be initiated using any HTTP client.

Define the webhook URL:

$url = "https://4534bh34-345j34d4-dfdfg435a3453.webhook.eus.azure-automation.net/webhooks?token=YdtdhY0RP2dsfsdS4204o5S2f9sERdfsdGMdfnAd%s3DadlxYDCQUBo%3d"

Create the body as a hashtable:

$body = @{
    location             = "Main Office"
    floornum            = "4"
    floorlabel          = "Fourth Floor"
    capacity            = "1"
    wheelchairaccessible = "true"
    officeid            = "345"
    cubicleid           = ""
    managers            = "someone@contoso.com,another@contoso.com"
    permissions         = "Editor (manage existing meetings),Approver"
    managementscope     = "All reservable resources (new and existing)"
    ticketid            = "34536"
    itemrequestid       = "10005161241"
}

Convert hashtable to JSON:

$jsonBody = $body | ConvertTo-Json -Depth 3

Send the POST request:

$response = Invoke-WebRequest -Uri $url -Method POST -Body $jsonBody -ContentType 'application/json'

Output the response:

$response.Content

.SYNOPSIS
    Automates the creation and configuration of reservable workspace resources in Exchange Online.

.DESCRIPTION
    This script ingests webhook data from a Freshservice Workflow Automator Workflow API call, and provisions a new room mailbox of type 'Workspace' in 
    Exchange Online, representing a reservable workspace (e.g., office or cubicle). It configures 
    location metadata, calendar processing rules, and access permissions based on the request payload.

    The script uses managed identity to securely authenticate with Azure and Exchange Online, 
    retrieves credentials from Azure Key Vault, and integrates with the Freshservice API to log 
    ticket updates and notify stakeholders of success or failure.

    Key operations include:
      - Parsing webhook JSON payloads
      - Creating Exchange Online room resource mailbox of type 'Workspace' (not available in Exchange Admin Center)
      - Setting location, capacity, and accessibility metadata
      - Assigning calendar permissions to managers and admin groups
      - Logging outcomes and updating Freshservice tickets

.AUTHOR
    Alex Gonzalez

.CREATED
    2024-02-15

.LAST MODIFIED
    2025-05-23

.VERSION
    2.0

.REQUIREMENTS
    - PowerShell 5.1 (7.1+ has existing issues with required modules as Microsoft botched compatiability [expected to be fixed ~2025])
    - ExchangeOnlineManagement module (any version)
    - Az.KeyVault module (v4.9.2) - highest version capabling of supporting current Az.Accounts dependency due to shenanigans with Microsoft breaking Azure Automation
    - Azure Managed Identity enabled with access to Exchange Online (App registration permission: Office 365 Exchange Online role with claim of Exchange.ManageAsApp) and Key Vault (IAM role of 'Key Vault Secret User assigned to Automation Account)
    - Freshservice API key stored in Azure Key Vault

.NOTES
    This script was designed to be used in the context of an Azure Automation runbook as a managed identity. It includes robust error handling and notification logic.
#>


Param(
     [parameter(Mandatory=$true)]
     [Object]$WebhookData
)

Import-Module ExchangeOnlineManagement
Import-Module Az.KeyVault -RequiredVersion "4.9.2"

# Validates request body is present and attempts to convert to PowerShell object.
if ($WebhookData.RequestBody) 
{
    try {
        $PayloadRequestBody = ConvertFrom-Json -InputObject $WebhookData.RequestBody
    }
    catch {
        throw "Unable to parse JSON in API request."
    }
}

# Set PS variables for basic office attributes.
$Office = $PayloadRequestBody.location
$FloorNum = $PayloadRequestBody.floornum
$FloorLabel = $PayloadRequestBody.floorlabel
$Capacity = $PayloadRequestBody.capacity
$WheelChairAccessible = $PayloadRequestBody.wheelchairaccessible
$OfficeId = $PayloadRequestBody.officeid
$CubicleId = $PayloadRequestBody.cubicleid
$Managers = $PayloadRequestBody.managers.Split(',').Trim()
$Permissions = $PayloadRequestBody.permissions.Split(',')
$ManagementScope = $PayloadRequestBody.managementscope
$TicketID = $PayloadRequestBody.ticketid
$ServiceRequestItemID = $PayloadRequestBody.itemrequestid -replace '[\[\]]', ''

# Sets organization name, domain, and Azure subscription ID.
$OrganizationName = 'OrganizationName'
$FSDomain = 'FreshserviceDomain'
$DomainName = 'DomainName'
$SubscriptionId = 'SubscriptionId'

# Sets Keyvault name, credential name, and administrative group variables for managing workspace resources.
$KeyvaultName = 'KeyvaultName'
$CredentialName = 'CredentialName'
$AdminGroup = 'AdminGroupEmail'

# Sets parameters for Get-AzKeyVaultSecret cmdlet to securely retrieve Mr. Automation's API creds for Freshservice API requests.
$KeyVaultParams = @{
Name = $CredentialName
VaultName = $KeyvaultName
AsPlainText = $true
}

# Set office and cubicle abbreviation values as well as resource type to be created.
$OfficePrefix = 'OF'
$CubiclePrefix = 'WS'
$ResourceType = 'Workspace'

# Initalize string variables for valid managers
$ValidManagers = $null
$InvalidManagers = $null

# Connect to Azure for retrieving credentials (requires managed identity to be enabled)
Connect-AzAccount -Subscription $SubscriptionId -Identity -Verbose

# Connect to Exchange Online using managed identity (requires managed identity to be enabled + Office 365 Exchange Online role with claim of Exchange.ManageAsApp)
Connect-ExchangeOnline -ManagedIdentity -Organization $DomainName -Verbose

# Sets API URLs while including unique ticket ID and service request item ID.
$FreshserviceCreatePrivateNoteUpdateURL = "https://$FSDomain/api/v2/tickets/$TicketID/notes"
$FreshserviceUpdateServiceRequestItemStatusURL = "https://$FSDomain/api/v2/tickets/$TicketID/requested_items/$ServiceRequestItemID"

# Sets header info for Freshservice API call. Retrieves Freshservice API key from Azure Key Vault and encodes using Base64 (requires )
$Headers = @{
    "Authorization" = ("Basic" + " " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('{0}:{1}' -f (Get-AzKeyVaultSecret @KeyVaultParams), $null))) )
    "Content-Type" = "application/json"
}

# Sets room list based on provided office name (requires pre-existing room lists to be defined - see New-ExchangeRoomList.ps1 to create)
if ($Office -eq "Office 1")
{
    $RoomList = "RoomListEmail1"
    $Building = "Building A"
    $Street = "1000 Fake Blvd, Ste 300"
    $City = "New York"
    $State = "NY"
    $Zipcode = "14534"
    $Country = 'United States'
}
elseif ($Office -eq "Office 2")
{
    $RoomList = "RoomListEmail2"
    $Building = "Building C"
    $Street = "443 Another Fake Blvd"
    $City = "San Francisco"
    $State = "CA"
    $Zipcode = "43249"
    $Country = 'United States'
}
elseif ($Office -eq "Office 3")
{
    $RoomList = "RoomListEmail3"
    $Building = "Building 200"
    $Street = "4234 Park Ave"
    $City = "Cincinatti"
    $State = "OH"
    $Zipcode = "34245"
    $Country = 'United States'
}

# Set username using office abbreviation code and office ID #.
if ($OfficeId)
{
    $IsOffice = $true
    $Username = ($Building.ToLower() + "-" + $OfficePrefix.ToLower() + "-" + $OfficeId)
    $DisplayName = ($OfficePrefix + " " + $OfficeId.ToUpper())
}
if ($CubicleId)
{
    $Username = ($Building.ToLower() + "-" + $CubiclePrefix.ToLower() + "-" + $CubicleId)
    $DisplayName = ($CubiclePrefix + " " + $CubicleId.ToUpper())
}

################################################################################################################  **Parameters for resource settings**  ###################################################################################################################################################

# Base parameters for Set-Place cmdlet. Indicates basic resource location information.
$SetPlaceParams = @{
    Identity = $Username
    Building = $Building
    Capacity = $Capacity
    Street = $Street
    City = $City
    State = $State
    PostalCode = $Zipcode
    CountryOrRegion = $Country
    Floor = $FloorNum
    FloorLabel = $FloorLabel
}
# Appends parameter to Set-Place cmdlet if resource is handicap accessible.
if ($WheelChairAccessible)
{
    # Parameters for Set-Place cmdlet.
    $SetPlaceParams += @{ IsWheelChairAccessible = $true }
}

# Parameters for Set-CalendarProcessing cmdlet.
$SetCalendarProcessingParams = @{
    Identity = $Username
    AutomateProcessing = "AutoAccept"
    AllowConflicts = $false
    AllowRecurringMeetings = $true
    EnforceCapacity = $true
    RemoveOldMeetingMessages = $true
    RemoveCanceledMeetings = $true
    Confirm = $false
}

# Parameters for Add-DistributionGroupMember cmdlet. Adds resource to room list (distribution list group) to allow resource to be found in Outlook Room Finder tool.
$AddDistributionGroupMemberParams = @{
    Identity = $RoomList
    Member = $Username
    Confirm = $false
}

# Parameters for New-Mailbox cmdlet. Creates room resource.
$NewMailboxParams = @{
    Name = $Username
    Room = $true
    Confirm = $false
}

# Parameters for Add-MailboxFolderPermission cmdlet. Adds editor rights to administrative groups.
$AddMailboxParams = @{
    Identity = $Username + ":\calendar"
    AccessRights = "Editor"
    Confirm = $false
}

# Parameters for Set-Mailbox cmdlet. Sets Display Name, Name, and type to Workspace.
$SetMailboxParams = @{
    Identity = $Username
    Type = $ResourceType
    Name = $DisplayName
    DisplayName = $DisplayName
    Confirm = $false
}

$SetUserParams = @{
    Identity = $Username
    Company =  $OrganizationName
    Confirm = $false
}

# Sets manager approval policy for office requests.
if ($IsOffice)
{
    $SetCalendarProcessingParams += @{
        AllRequestInPolicy = $true
        AllBookInPolicy = $false
        ForwardRequestsToDelegates = $true
        TentativePendingApproval = $true
        AddNewRequestsTentatively = $true
    }
}

# Sets management scope of resources (single vs all new and existing)
if ($ManagementScope -eq "Single reservable resource (this resource only)")
{
    $SingleResourceScope = $true
}
elseif ($ManagementScope -eq "All reservable resources (new and existing)")
{
    $AllResourceScope = $true
}

################################################################################################################  ** Manager and calendar permission settings for resource settings **  ##################################################################################################################


# Set parameters for Set-CalendarProcessing cmdlet based on if manager is provided.
if ($Managers)
{  
    $OldPref = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'

    foreach ($Manager in $Managers)
    {
        Write-Output "Looping through managers for validation. Current manager: $Manager"
        try 
        {
            # Validate manager email address before setting processing rules.
            Get-EXOMailbox -Identity $Manager

            # Concatenate valid manager to string.
            $ValidManagers = $ValidManagers + $Manager + ","
            
            # Loop through each permission provided in request
            foreach ($Permission in $Permissions)
            {
                # Set parameters for adding editor permissions to resource calendar.
                if ($Permission -eq "Editor (manage existing meetings)")
                {
                    $EditorRights = $true
                    Write-Output "Editor permissions assigned to $Manager"
                }

                # Add manager approver to resource.
                if ($Permission -eq "Approver")
                {
                    # Set manager rights boolean to true.
                    $ApproverRights = $true
                    Write-Output "Approver permissions assigned to $Manager"
                }
            }
        } 
        catch 
        {
            # Write error output to stream.
            Write-Error "Unable to find manager $Manager in Exchange. Skipping manager assignment"

            # Concatenate invalid manager to string.
            $InvalidManagers= $InvalidManagers + $Manager + ","
        }
    }
    $ErrorActionPreference = $OldPref
}


################################################################################################################  **Runs cmdlets to set various settings defined in "Parameters for resource settings" and API calls**  #####################################################################################

# Sets API request body request based on success, failure, or warnings.
$NewPrivateNoteSuccessBody = '{ "body":"<div>The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' has successfully created. <br><br> Please allow up to 24 hours for the resource to appear in Outlook Room Finder.</div>", "private":true }'
$NewPrivateNoteFailureBody = '{ "body":"<div>The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' has failed to create. <br><br> Please reach out to your systems administrator for further assistance. Do <b>NOT</b> re-submit this request.</div>", "private":true }'
$NewPrivateNoteResourceExistsBody = '{ "body":"The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' already exists. <br><br> Please check the information provided and try again by creating a new service request ticket.</div>", "private":true }'
if ($InvalidManagers) { $InvalidManagerBody = '{ "body":"<div>The manager(s) ' + '<b>' + $InvalidManagers.Trim(',') + '</b>' + ' do not contain valid email address(es).<br><br> Please reach out to your systems administrator for further assistance. Do <b>NOT</b> re-submit this request.</div>", "private":true }' }
$UpdateRequestedItemStatusCancelledBody = '{ "stage":3 }'
$UpdateRequestedItemStatusFulfilledBody = '{ "stage":4 }'

# Check if identity exists before attempting operations. If no results are returned, proceed.

try 
{
    $OldPref = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'

    # Attempt to retrieve existing identity.
    $PreMailboxCheck = Get-EXOMailbox -Identity $Username
}
catch 
{
    # Create resource mailbox.
    New-Mailbox @NewMailboxParams

    # Set Company Name attribute associated with resource.
    Set-User @SetUserParams
    
    # Set mailbox Display Name and type to Workspace.
    Set-Mailbox @SetMailboxParams

    # Set 30 second timer to allow resource to propagate prior to setting other resource values.
    Start-Sleep -Seconds 10

    # Set workspace details for location capacity, country, floor number, floor label, and wheelchair accessability.
    Set-Place @SetPlaceParams

    # Add Workspace as member to room list based on desginated Room List for an office.
    Add-DistributionGroupMember @AddDistributionGroupMemberParams

    # Set resource calendar processing rules.
    Set-CalendarProcessing @SetCalendarProcessingParams

    $ErrorActionPreference = $OldPref


    # Assigns appropriate delegate and or editor permissions to this resource as well as any future and existing resources.
    if ($AllResourceScope) 
    {
        # Assigns "approver" delegate and calendar editor permissions to all resource management group.
        if ($EditorRights -and $ApproverRights -and $ValidManagers)
        {
            foreach ($Manager in $ValidManagers.Split(','))
            {
                try 
                { 
                    # Assign editor rights to mail-enabled security group.
                    Add-DistributionGroupMember -Identity $AdminGroup -Member $Manager 
                } 
                catch { $_ }
            }

            try 
            {
                # Assigns calendar editor permissions to admin group.
                Add-MailboxFolderPermission @AddMailboxParams -User $AdminGroup 
            }
            catch { $_ }

            try
            {
                # Updates calendar processing to include administrative group as a delegate.
                Set-CalendarProcessing -Identity $Username -ResourceDelegates $AdminGroup -Confirm:$false 
            }
            catch { $_}
        }


        # Sets manager permissions on resource mailbox if flag for approver rights are provided in the initial request.
        elseif ($EditorRights -and $ValidManagers)
        {
            foreach ($Manager in $ValidManagers.Split(','))
            {
                try 
                { 
                    # Assign editor rights to mail-enabled security group.
                    Add-DistributionGroupMember -Identity $AdminGroup -Member $Manager 
                } 
                catch { $_ }
            }

            try 
            {
                # Assigns calendar editor permissions to admin group.
                Add-MailboxFolderPermission @AddMailboxParams -User $AdminGroup 
            }
            catch { $_ }
        }

        # Check if manager approver flag was added in request.
        elseif ($ApproverRights -and $ValidManagers)
        {
            try
            {
                # Updates calendar processing to include administrative group as a delegate.
                Set-CalendarProcessing -Identity $Username -ResourceDelegates $AdminGroup -Confirm:$false 
            }
            catch { $_ }

            try
            {
                # Removes unintended editor rights to calendar automatically applied when using the Set-CalendarProcessing cmdlet to add resource delegates.
                Remove-MailboxFolderPermission -Identity "$($Username):\calendar" -User $AdminGroup -Confirm:$false 
            }
            catch { $_ }
        }
    }
}

# Assigns appropriate delegate and or editor permissions to this resource onlu for all managers. 
elseif ($SingleResourceScope)
{
    # Assigns "approver" delegate and calendar editor permissions to individuals.
    if ($EditorRights -and $ApproverRights -and $ValidManagers)
    {
        foreach ($Manager in $ValidManagers.Split(','))
        {
            try 
            {
                # Assign editor rights to indvidual managers on this resource only.
                Add-MailboxFolderPermission @AddMailboxParams -User $Manager
            }
            catch { $_ }
        }
        try 
        {
            # Updates calendar processing to include specific managers as delegate to this resource only.
            Set-CalendarProcessing -Identity $Username -ResourceDelegates $ValidManagers.Trim(',') -Confirm:$false
        }
        catch { $_ }
    }

    # Sets manager permissions on resource mailbox if flag for approver rights are provided in the initial request.
    elseif ($EditorRights -and $ValidManagers)
    {
        foreach ($Manager in $ValidManagers.Split(','))
        {
            try 
            {
                # Assign editor rights to indvidual managers on this resource only.
                Add-MailboxFolderPermission @AddMailboxParams -User $Manager
            }
            catch { $_ }
        }
    }

    # Check if manager approver flag was added in request.
    elseif ($ApproverRights -and $ValidManagers)
    {
        try 
        {
            # Updates calendar processing to include specific managers as delegate to this resource only.
            Set-CalendarProcessing -Identity $Username -ResourceDelegates $ValidManagers.Trim(',') -Confirm:$false
        }
        catch { $_ }

        try
        {
            $ManagerList = $ValidManagers.Trim(',').Split(',') | Where-Object { $_ -ne "" }

            foreach ($Manager in $ManagerList) 
            {
                try
                {
                    # Removes unintended editor rights to calendar automatically applied when using the Set-CalendarProcessing cmdlet to add resource delegates.
                    Remove-MailboxFolderPermission -Identity "$($Username):\calendar" -User $Manager -Confirm:$false 
                }
                catch { $_ }
            }
        }
        catch { $_ }
    }
}

################################################################################################################  **Post resource mailbox creation check**  #############################################################################################################################################

if (-not $PreMailboxCheck)
{
    try 
    {
        $OldPref = $ErrorActionPreference
        $ErrorActionPreference = 'Stop'

        # Verify post mailbox creation.
        $PostMailboxCheck = Get-EXOMailbox -Identity $Username

        # Output success message to stream
        Write-Output "The resource '$($Username)' created successfully."

        $ErrorActionPreference = $OldPref
        
        if ($PostMailboxCheck)
        {
            try
            {
                # Create private note with success status, and update requested item status to 'Fullfilled'.
                Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteSuccessBody -UseBasicParsing
                Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusFulfilledBody -UseBasicParsing
            }
            catch { Write-Error "Unable to update ticket status indicating that '$($Username)' created successfully with requested item status of 'Fulfulled'." }
        }
    }
    catch 
    {
        # Output error to stream.
        Write-Error "The resource '$($Username)' failed to create."

        try
        {
            # Create private note with failure status and update requested item status to 'Cancelled'.
            Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteFailureBody -UseBasicParsing
            Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusCancelledBody -UseBasicParsing
        }
        catch { Write-Error "Unable to update ticket status indicating that '$($Username)' failed to create with requested item status of 'Cancelled'."}
    }
}

if ($InvalidManagers)
{
    try
    {
        # Create private note indicating that the resource manager was not applied due to an invalid email address.
        Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $InvalidManagerBody -UseBasicParsing
    }
    catch { Write-Error "Unable to update ticket status indicating that '$($InvalidManagers)' are invalid." }

}

if ($PreMailboxCheck)
{
    # Output error to stream.
    Write-Error "The resource $Username already exists."

    try
    {
        # Create private note indicating that resource already exists and update requested item status to 'Cancelled'.
        Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteResourceExistsBody -UseBasicParsing
        Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusCancelledBody -UseBasicParsing
    }
    catch { Write-Error "Unable to update ticket status indicating the resource '$($Username)' already exists with requested item status of 'Cancelled'." }

}

# Disconnect from Exchange Online session.
Disconnect-ExchangeOnline -Confirm:$false


#######################################################################################################################################################################################################################################################################################################
