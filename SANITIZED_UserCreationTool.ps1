# Define log file path
$LogFile = "C:\UserCreationTool\UserProvisioningLog.txt"

# Function to write to log
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $logEntry
    Write-Host $logEntry
}

# Ensure Microsoft Graph Groups module is available
Import-Module Microsoft.Graph.Groups
Connect-MgGraph -NoWelcome
Write-Log "Connected to Microsoft Graph."

# Prompt for user input
$FirstName = Read-Host "Enter First Name"
$LastName = Read-Host "Enter Last Name"
$Username = Read-Host "Enter Username (e.g., jdoe)"
$Password = Read-Host "Enter Password" -AsSecureString
$OU = Read-Host "Enter OU path (e.g., OU=Users,DC=YOUR,DC=DOMAIN)"
$ReferenceUsername = Read-Host "Enter reference username to copy group memberships from"

# Check if OU is valid
try {
    Get-ADOrganizationalUnit -Identity $OU -ErrorAction Stop
    Write-Log "Validated OU: $OU"
} catch {
    Write-Log "Invalid OU: $OU" "ERROR"
    Write-Error "The OU '$OU' does not exist. Please check the path and try again."
    return
}

# Construct UPN and routing address
$UPN = "$Username@kyourdomain.com"
$RemoteRouting = "$Username@yourdomain.mail.onmicrosoft.com"

# Get reference user's local AD group memberships
$ReferenceGroups = Get-ADUser -Identity $ReferenceUsername -Properties MemberOf | Select-Object -ExpandProperty MemberOf
Write-Log "Retrieved local AD groups for reference user: $ReferenceUsername"

# Create AD User
try {
    New-ADUser -Name "$FirstName $LastName" `
        -GivenName $FirstName -Surname $LastName `
        -SamAccountName $Username -UserPrincipalName $UPN `
        -AccountPassword $Password -Path $OU -Enabled $true
    Write-Log "Created AD user: $Username"
} catch {
    Write-Log "Failed to create AD user: $Username. Error: $_" "ERROR"
}

# Add to same groups as reference user (local AD only)
foreach ($groupDN in $ReferenceGroups) {
    try {
        $group = Get-ADGroup -Identity $groupDN -ErrorAction Stop
        Add-ADGroupMember -Identity $group.Name -Members $Username
        Write-Log "Added $Username to local AD group: $($group.Name)"
    } catch {
        Write-Log "Skipping cloud-only or invalid group: $groupDN" "WARNING"
    }
}


# Add to hybrid groups (on-prem AD groups that sync to Azure AD)
foreach ($groupDN in $ReferenceGroups) {
    # Skip cloud-only groups that use GUID-based names
    if ($groupDN -like "*Group_*") {
        Write-Log "Skipping cloud-only group: $groupDN" "WARNING"
        continue
    }

    try {
        $group = Get-ADGroup -Identity $groupDN -Properties mail
        Add-ADGroupMember -Identity $group.Name -Members $Username
        Write-Log "Added $Username to hybrid AD group: $($group.Name)"
    } catch {
        Write-Log "Failed to process group: $groupDN. Error: $_" "WARNING"
   }

}

# Add user to AD writeback groups (Group_<GUID>) if they exist in AD
foreach ($groupDN in $ReferenceGroups) {
    if ($groupDN -like "*Group_*") {
        try {
            # Extract CN from DN
            $groupCN = ($groupDN -split ",")[0] -replace "CN=", ""

            # Search for group by Name using -Filter
            $writebackGroup = Get-ADGroup -Filter "Name -eq '$groupCN'" -ErrorAction Stop

            if ($writebackGroup) {
                Add-ADGroupMember -Identity $writebackGroup.DistinguishedName -Members $Username
                Write-Log "Added $Username to AD writeback group: $($writebackGroup.Name)"
            } else {
                Write-Log "Writeback group not found in AD: $groupCN" "WARNING"
            }
        } catch {
            Write-Log "Writeback group not found or failed to add: $groupCN. Error: $_" "WARNING"
        }
    }
}

# Connect to on-prem Exchange server
$ExchangeServer = "yourexchangeserver"
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchangeServer/PowerShell/" -Authentication Kerberos
    Import-PSSession $Session -DisableNameChecking
    Write-Log "Connected to Exchange server: $ExchangeServer"
} catch {
    Write-Log "Failed to connect to Exchange server: $ExchangeServer. Error: $_" "ERROR"
}

# Enable Remote Mailbox
try {
    Enable-RemoteMailbox -Identity $Username -RemoteRoutingAddress $RemoteRouting
    Write-Log "Enabled remote mailbox for: $Username"
} catch {
    Write-Log "Failed to enable remote mailbox for: $Username. Error: $_" "ERROR"
}

# Clean up Exchange session
Remove-PSSession $Session
Write-Log "Removed Exchange session."

#Sync AD
Start-AdSyncSyncCycle -PolicyType Delta

# Wait 10 minutes before proceeding
Write-Log "Waiting 10 minutes before proceeding to Teams provisioning..."
Start-Sleep -Seconds 600

# Retry logic to get Azure AD user object
$NewUserAAD = $null
$retryCount = 0
$maxRetries = 10
$retryDelaySeconds = 60

do {
    try {
        $NewUserAAD = Get-MgUser -UserId $UPN
        if ($NewUserAAD) {
            Write-Log "Successfully retrieved Azure AD user object for: $UPN"
            break
        } else {
            Write-Log "Azure AD user not found yet: $UPN. Retrying in $retryDelaySeconds seconds..." "WARNING"
        }
    } catch {
        Write-Log "Error retrieving Azure AD user: $UPN. Retrying... Error: $_" "WARNING"
    }

    Start-Sleep -Seconds $retryDelaySeconds
    $retryCount++
} while ($retryCount -lt $maxRetries)

if (-not $NewUserAAD) {
    Write-Log "Azure AD user not found after $maxRetries retries: $UPN" "ERROR"
    return
}


# Connect to Microsoft Teams
Connect-MicrosoftTeams
Write-Log "Connected to Microsoft Teams."

# Get all Teams the source user is a member of
$teams = Get-Team | Where-Object {
    Get-TeamUser -GroupId $_.GroupId | Where-Object { $_.User -eq $ReferenceUsername }
}
Write-Log "Retrieved Teams memberships for reference user: $ReferenceUsername"

# Add the target user to each team
foreach ($team in $teams) {
    try {
        Add-TeamUser -GroupId $team.GroupId -User $Username
        Write-Log "Added $Username to team: $($team.DisplayName)"
    } catch {
        Write-Log "Failed to add $Username to team: $($team.DisplayName). Error: $_" "WARNING"
    }
}

# Cloud group membership provisioning
try {
    $ReferenceUserAAD = Get-MgUser -UserId "$ReferenceUsername@yourdomain.com"
    $NewUserAAD = Get-MgUser -UserId $UPN
    $CloudGroups = Get-MgUserMemberOf -UserId $ReferenceUserAAD.Id | Where-Object { $_.ODataType -eq "#microsoft.graph.group" }
    Write-Log "Retrieved Azure AD group memberships for reference user."

    foreach ($group in $CloudGroups) {
        try {
            Add-MgGroupMember -GroupId $group.Id -DirectoryObjectId $NewUserAAD.Id
            Write-Log "Added $Username to Azure AD group: $($group.AdditionalProperties["displayName"])"
        } catch {
            Write-Log "Failed to add $Username to Azure AD group: $($group.Id). Error: $_" "WARNING"
        }
    }
} catch {
    Write-Log "Failed to retrieve or assign Azure AD group memberships. Error: $_" "ERROR"
}

# Assign Microsoft 365 license
$SkuId = "<your license sku id>"
try {
    Set-MgUser -UserId $NewUserAAD.Id -UsageLocation "US"
    Set-MgUserLicense -UserId $NewUserAAD.Id -AddLicenses @{SkuId=$SkuId} -RemoveLicenses @()
    Write-Log "Assigned license ($SkuId) to user: $Username"
} catch {
    Write-Log "Failed to assign license to user: $Username. Error: $_" "ERROR"
}


