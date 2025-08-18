Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "User Provisioning Tool"
$form.Size = [System.Drawing.Size]::new(520, 520)
$form.StartPosition = "CenterScreen"

# Create input fields
$fields = @{}
$fieldLabels = @("First Name", "Last Name", "Username", "Password", "OU Path", "Reference Username")

$i = 0
foreach ($labelText in $fieldLabels) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = [System.Drawing.Point]::new(20, (30 + ($i * 40)))
    $label.Size = [System.Drawing.Size]::new(120, 20)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = [System.Drawing.Point]::new(150, (30 + ($i * 40)))
    $textbox.Size = [System.Drawing.Size]::new(320, 20)
    if ($labelText -eq "Password") { $textbox.UseSystemPasswordChar = $true }
    $form.Controls.Add($textbox)

    $fields[$labelText] = $textbox
    $i++
}

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = [System.Drawing.Point]::new(20, 300)
$progressBar.Size = [System.Drawing.Size]::new(450, 20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$form.Controls.Add($progressBar)

# Progress label
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = [System.Drawing.Point]::new(20, 330)
$progressLabel.Size = [System.Drawing.Size]::new(450, 20)
$progressLabel.Text = "Progress: Waiting to start..."
$form.Controls.Add($progressLabel)

# Provision button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Provision User"
$button.Location = [System.Drawing.Point]::new(200, 370)
$button.Size = [System.Drawing.Size]::new(120, 30)
$form.Controls.Add($button)

# Log file path
$LogFile = "C:\UserCreationTool\UserProvisioningLog.txt"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $logEntry
    if ($Level -eq "ERROR" -or $Level -eq "WARNING") {
        [System.Windows.Forms.MessageBox]::Show($logEntry)
    }
}

function Update-Progress {
    param (
        [int]$Percent,
        [string]$Message
    )
    $progressBar.Value = $Percent
    $progressLabel.Text = "Progress: $Message"
    $form.Refresh()
}

function Validate-Inputs {
    foreach ($key in $fields.Keys) {
        if ([string]::IsNullOrWhiteSpace($fields[$key].Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in the '$key' field.")
            return $false
        }
    }
    return $true
}

function Validate-UserDoesNotExist {
    param ([string]$Username)
    try {
        $existingUser = Get-ADUser -Identity $Username -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("User '$Username' already exists in Active Directory.")
        return $false
    } catch {
        return $true
    }
}

# Button click event
$button.Add_Click({
    if (-not (Validate-Inputs)) { return }

    $FirstName = $fields["First Name"].Text
    $LastName = $fields["Last Name"].Text
    $Username = $fields["Username"].Text
    $Password = ConvertTo-SecureString $fields["Password"].Text -AsPlainText -Force
    $OU = $fields["OU Path"].Text
    $ReferenceUsername = $fields["Reference Username"].Text
    $UPN = "$Username@yourdomain.com"
    $RemoteRouting = "$Username@yourdomain.mail.onmicrosoft.com"

    if (-not (Validate-UserDoesNotExist -Username $Username)) { return }

    try {
        Update-Progress 5 "Connecting to Microsoft Graph..."
        Import-Module Microsoft.Graph.Groups
        Connect-MgGraph -NoWelcome
        Write-Log "Connected to Microsoft Graph."

        Update-Progress 10 "Validating OU..."
        Get-ADOrganizationalUnit -Identity $OU -ErrorAction Stop
        Write-Log "Validated OU: $OU"

        Update-Progress 15 "Retrieving reference user groups..."
        $ReferenceGroups = Get-ADUser -Identity $ReferenceUsername -Properties MemberOf | Select-Object -ExpandProperty MemberOf
        Write-Log "Retrieved local AD groups for reference user: $ReferenceUsername"

        Update-Progress 20 "Creating AD user..."
        New-ADUser -Name "$FirstName $LastName" `
            -GivenName $FirstName -Surname $LastName `
            -SamAccountName $Username -UserPrincipalName $UPN `
            -AccountPassword $Password -Path $OU -Enabled $true
        Write-Log "Created AD user: $Username"

        Update-Progress 30 "Adding to local AD groups..."
        foreach ($groupDN in $ReferenceGroups) {
            try {
                $group = Get-ADGroup -Identity $groupDN -ErrorAction Stop
                Add-ADGroupMember -Identity $group.Name -Members $Username
                Write-Log "Added $Username to local AD group: $($group.Name)"
            } catch {
                Write-Log "Skipping cloud-only or invalid group: $groupDN" "WARNING"
            }
        }

        Update-Progress 40 "Enabling remote mailbox..."
        $ExchangeServer = "yourexchange"
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchangeServer/PowerShell/" -Authentication Kerberos
        Import-PSSession $Session -DisableNameChecking
        Write-Log "Connected to Exchange server: $ExchangeServer"

        Enable-RemoteMailbox -Identity $Username -RemoteRoutingAddress $RemoteRouting
        Write-Log "Enabled remote mailbox for: $Username"
        Remove-PSSession $Session
        Write-Log "Removed Exchange session."

        Update-Progress 50 "Starting AD sync..."
        Start-AdSyncSyncCycle -PolicyType Delta
        Write-Log "Waiting 5 minutes before proceeding to Teams provisioning..."
        Start-Sleep -Seconds 300

        Update-Progress 60 "Checking Azure AD user..."
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
                    Write-Log "Azure AD user not found yet: $UPN. Retrying..." "WARNING"
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

        Update-Progress 70 "Connecting to Microsoft Teams..."
        try {
            Connect-MicrosoftTeams -ErrorAction Stop
            Write-Log "Connected to Microsoft Teams."
        } catch {
            Write-Log "Failed to connect to Microsoft Teams. Error: $_" "ERROR"
            return
        }


        $teams = Get-Team | Where-Object {
            Get-TeamUser -GroupId $_.GroupId | Where-Object { $_.User -eq $ReferenceUsername }
        }
        Write-Log "Retrieved Teams memberships for reference user: $ReferenceUsername"

        foreach ($team in $teams) {
            try {
                Add-TeamUser -GroupId $team.GroupId -User $Username
                Write-Log "Added $Username to team: $($team.DisplayName)"
            } catch {
                Write-Log "Failed to add $Username to team: $($team.DisplayName). Error: $_" "WARNING"
            }
        }

        Update-Progress 85 "Adding to Azure AD groups..."
        $ReferenceUserAAD = Get-MgUser -UserId "$ReferenceUsername@yourdomain.com"
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

        Update-Progress 95 "Assigning license..."
        $SkuId = "18181a46-0d4e-45cd-891e-60aabd171b4e"
        Set-MgUser -UserId $NewUserAAD.Id -UsageLocation "US"
        Set-MgUserLicense -UserId $NewUserAAD.Id -AddLicenses @{SkuId=$SkuId} -RemoveLicenses @()
        Write-Log "Assigned license ($SkuId) to user: $Username"

        Update-Progress 100 "Provisioning complete!"
        Write-Log "Provisioning complete for $Username"
    } catch {
        Write-Log "Provisioning failed. Error: $_" "ERROR"
    }
})

# Show form
$form.ShowDialog()

