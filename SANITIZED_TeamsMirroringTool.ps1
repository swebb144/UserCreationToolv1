# Prompt for source and target user UPNs
$sourceUser = Read-Host "Enter the source user's username"
$targetUser = Read-Host "Enter the target user's username"
$sourceUser = "$sourceUser@<yourdomain>.com"
$targetUser = "$targetUser@<yourdomain>.com"

# Define log file path
$logFile = "$PSScriptRoot\TeamsMirrorLog.log"

# Clear previous log file if exists
if (Test-Path $logFile) {
    Remove-Item $logFile
}

# Connect to Microsoft Teams
try {
    $null = Get-Team
} catch {
    Write-Host "Connecting to Microsoft Teams..."
    Connect-MicrosoftTeams
}

# Get all teams
$allTeams = Get-Team

foreach ($team in $allTeams) {
    try {
        $members = Get-TeamUser -GroupId $team.GroupId
        $isMember = $members | Where-Object { $_.User -eq $sourceUser }

        if ($isMember) {
            try {
                Add-TeamUser -GroupId $team.GroupId -User $targetUser
                $successMessage = "[$(Get-Date)] Added $targetUser to team: $($team.DisplayName)"
                Write-Host $successMessage
                Add-Content -Path $logFile -Value $successMessage
            } catch {
                $errorMessage = "[$(Get-Date)] Failed to add $targetUser to team: $($team.DisplayName). Error: $($_.Exception.Message)"
                Write-Host $errorMessage
                Add-Content -Path $logFile -Value $errorMessage
            }
        }
    } catch {
        $errorMessage = "[$(Get-Date)] Failed to retrieve members for team: $($team.DisplayName). Error: $($_.Exception.Message)"
        Write-Host $errorMessage
        Add-Content -Path $logFile -Value $errorMessage
    }
}

Write-Host "Script completed. Check '$logFile' for full log details."
