# PowerShell script to install Office and log the process

# Log file path
$logFile = "C:\OEM\setup_evkey.log"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Function to write to log file
function Write-Log {
    param ($Message)
    "$timestamp $Message" | Out-File -FilePath $logFile -Append
}

Write-Log "Starting InstallEVKey.ps1"
Start-Sleep -Seconds 30
Write-Log "Waiting a bit to make sure the system is ready"

# Check if EVKey is already installed
Write-Log "Checking if EVKey64 is already installed"
if (Test-Path "C:\Program Files\EVKey\EVKey64.exe") {
    Write-Log "EVKey is already installed. Ensuring success marker exists."
    try {
        New-Item -Path "C:\OEM\success" -ItemType File -Force | Out-Null
        Write-Log "Success file created (already installed)."
    } catch {
        Write-Log "Failed to create success file."
    }
    Write-Log "Exiting script as EVKey is already installed."
    exit 0
}

# Download EVKey64.exe (Portable)
$filePath = "C:\Program Files\EVKey\EVKey64.exe"
$zipFilePath = "C:\Program Files\EVKey\EVKey64.zip"
$folderPath = "C:\Program Files\EVKey"

Write-Log "Downloading EVKey..."
try {
    New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
    Invoke-WebRequest -Uri "https://github.com/lamquangminh/EVKey/releases/download/Release/EVKey.zip" -OutFile $zipFilePath -ErrorAction Stop
    Write-Log "Downloaded EVKey64.zip from primary URL."
    Write-Log "Extracting."
    Expand-Archive -Path $zipFilePath -DestinationPath $folderPath -ErrorAction Stop 
} catch {
    Write-Log "Download from fallback URL [Github]"
    try {
        Invoke-WebRequest -Uri "[REPLACEMENT]" -OutFile $filePath -ErrorAction Stop  # I will upload this on Github and use Github download link instead
        Write-Log "Downloaded EVKey64.exe from fallback URL."
    } catch {
        Write-Log "Failed to download or extract from fallback URL."
    }
}
