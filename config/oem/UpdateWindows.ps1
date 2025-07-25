# Update Windows and Office
# Can be run through FreeRDP with the xfreerdp parameter `/app:program:powershell.exe,cmd:"-ExecutionPolicy Bypass -File C:\OEM\update.ps1"`

# Start logging
Start-Transcript -Path "C:\OEM\update.log" -Append

# Check if running via FreeRDP
$isFreeRDP = $false
$process = Get-CimInstance Win32_Process -Filter "ProcessId = $PID"
$commandLine = $process.CommandLine
if ($commandLine -match "xfreerdp|xfreerdp3|freerdp|flatpak") {
    $isFreeRDP = $true
}

try {
    # Check if running as Administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Host "This script requires administrative privileges. Attempting to elevate..."
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }

    # Set module installation path
    $modulePath = "C:\Program Files\WindowsPowerShell\Modules"
    if (-not (Test-Path $modulePath)) {
        New-Item -Path $modulePath -ItemType Directory -Force | Out-Null
    }
    $env:PSModulePath = $modulePath # Override default module path

    # Check if PSWindowsUpdate module is installed, install if not
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Host "Installing PSWindowsUpdate module to $modulePath..."
        try {
            Install-Module -Name PSWindowsUpdate -Force -Scope AllUsers -ErrorAction Stop
        } catch {
            Write-Host "Failed to install PSWindowsUpdate module: $_"
            exit 1
        }
    }

    # Import PSWindowsUpdate module
    try {
        Import-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
    } catch {
        Write-Host "Failed to import PSWindowsUpdate module: $_"
        exit 1
    }

    # Configure update settings
    Write-Host "Checking for Windows and Office updates..."
    try {
        $updates = Get-WUList -MicrosoftUpdate -ErrorAction Stop
    } catch {
        Write-Host "Failed to retrieve updates: $_"
        exit 1
    }

    if ($updates.Count -eq 0) {
        Write-Host "No updates available."
        exit
    }

    # Display available updates
    Write-Host "`nFound $($updates.Count) update(s):"
    $updates | ForEach-Object { Write-Host "- $($_.Title)" }

    # Install updates
    Write-Host "`nInstalling updates..."
    try {
        Install-WindowsUpdate -AcceptAll -IgnoreReboot -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Host "Failed to install updates: $_"
        exit 1
    }

    # Check if reboot is required
    $rebootRequired = Get-WURebootStatus -Silent

    if ($rebootRequired) {
        Write-Host "`nA system reboot is required to complete the update installation."
        $response = Read-Host "Would you like to reboot now? (y/N)"
        if ($response -eq 'y' -or $response -eq 'Y') {
            Write-Host "Rebooting system..."
            Restart-Computer -Force
        } else {
            Write-Host "Please reboot your system manually or run the update script again and select 'Y' to complete the update process."
        }
    } else {
        Write-Host "`nUpdate installation completed. No reboot required."
    }

    # Pause only for non-FreeRDP sessions
    if (-not $isFreeRDP) {
        Write-Host "`nPress any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
} finally {
    # Clean up unwanted WindowsPowerShell folder only if empty
    $unwantedFolder = "\\tsclient\home\WindowsPowerShell"
    if (Test-Path $unwantedFolder) {
        $files = Get-ChildItem -Path $unwantedFolder -Recurse -File -ErrorAction SilentlyContinue
        if (-not $files) {
            Remove-Item -Path $unwantedFolder -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Removed empty or folder-only unwanted folder: $unwantedFolder"
        } else {
            Write-Host "Folder $unwantedFolder contains files. Skipping deletion."
        }
    }

    # Ensure clean exit for FreeRDP
    if ($isFreeRDP) {
        Write-Host "Terminating PowerShell session for FreeRDP."
        exit 0
    }
}

Stop-Transcript