# trigger_workflow.ps1
# Script to trigger the GitHub Actions Website Monitor workflow remotely.
#
# SETUP:
# Set MONITOR_GITHUB_PAT in your user/machine environment before running:
#   setx MONITOR_GITHUB_PAT "ghp_yourPersonalAccessTokenHere"
# Optionally override the repository:
#   setx MONITOR_GITHUB_REPO "your-username/Monitor-App"

$githubToken = [Environment]::GetEnvironmentVariable("MONITOR_GITHUB_PAT", "User")
if (-not $githubToken) {
    $githubToken = [Environment]::GetEnvironmentVariable("MONITOR_GITHUB_PAT", "Machine")
}
if (-not $githubToken) {
    $githubToken = $env:MONITOR_GITHUB_PAT
}

if (-not $githubToken) {
    throw "MONITOR_GITHUB_PAT is not set. Create a GitHub personal access token with workflow permissions and set it as an environment variable."
}

# Target repository: uses MONITOR_GITHUB_REPO if set, otherwise placeholder
$githubRepo = [Environment]::GetEnvironmentVariable("MONITOR_GITHUB_REPO", "User")
if (-not $githubRepo) {
    $githubRepo = $env:MONITOR_GITHUB_REPO
}
if (-not $githubRepo) {
    # Replace with your GitHub 'username/repo-name'
    $githubRepo = "YOUR_GITHUB_USERNAME/Monitor-App"
}

$headers = @{
    "Accept"               = "application/vnd.github+json"
    "Authorization"        = "Bearer $githubToken"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent"           = "PowerShell"
}

$body = @{
    ref = "main"
} | ConvertTo-Json

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logPath   = Join-Path $ScriptDir "trigger_log.txt"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

try {
    $response = Invoke-RestMethod -Uri "https://api.github.com/repos/$githubRepo/actions/workflows/website-monitor.yml/dispatches" -Method Post -Headers $headers -Body $body -ContentType "application/json"
    $logMsg = "[$timestamp] ? Successfully triggered GitHub Actions workflow for $githubRepo."
    Write-Output $logMsg
    Add-Content -Path $logPath -Value $logMsg
} catch {
    $logMsg = "[$timestamp] ? Failed to trigger workflow: $_"
    Write-Error $logMsg
    Add-Content -Path $logPath -Value $logMsg
}
