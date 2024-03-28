<#
.SYNOPSIS
This script resets Microsoft Teams.

.DESCRIPTION
This script stops the Microsoft Teams (new) process and clears its cache directory.

.PARAMETER None
This script does not accept any parameters.

.EXAMPLE
.\reset_teams.ps1
Resets Microsoft Teams.

.NOTES
File Name: reset_teams.ps1
Author: Matisse Vuylsteke
Date Created: 28/03/2024
#>

# Check if any process with a name like "Teams" is running
$teamsProcess = Get-Process | Where-Object { $_.ProcessName -like "*Teams*" }

if ($teamsProcess) {
    Write-Host "Resetting Teams..."
    # Quit Teams
    Stop-Process -Name $teamsProcess.ProcessName -Force

    # Get Teams local cache directory path
    $teamsCachePath = "$env:userprofile\appdata\local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams"

    # Stop msedgewebview2.exe process if it's running
    $edgeWebViewProcess = Get-Process -Name msedgewebview2 -ErrorAction SilentlyContinue
    if ($edgeWebViewProcess) {
        Stop-Process -Name msedgewebview2 -Force
    }

    # Delete all files and folders in the directory
    $cacheItems = Get-ChildItem -Path $teamsCachePath -Recurse
    foreach ($item in $cacheItems) {
        try {
            Remove-Item -Path $item.FullName -Force -Recurse -ErrorAction Stop
        } catch {
            if (-not $_.Exception.Message.Contains("because it does not exist.")) {
                Write-Output "Failed to delete $($item.FullName): $_"
            }
        }
    }
    Write-Host "Successfully reset Teams." -ForegroundColor Green
} else {
    Write-Output "Teams is not running."
}
