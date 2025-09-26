# PowerShell script to install Office and log the process

# Log file path
$logFile = "C:\OEM\setup_office.log"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Function to write to log file
function Write-Log {
    param ($Message)
    "$timestamp $Message" | Out-File -FilePath $logFile -Append
}

Write-Log "Starting InstallOffice.ps1"
Start-Sleep -Seconds 30
Write-Log "Waiting a bit to make sure the system is ready"

# Check if MS Office is already installed
Write-Log "Checking if MS Office is already installed"
if (Test-Path "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") {
    Write-Log "MS Office is already installed. Ensuring success marker exists."
    try {
        New-Item -Path "C:\OEM\success" -ItemType File -Force | Out-Null
        Write-Log "Success file created (already installed)."
    } catch {
        Write-Log "Failed to create success file."
    }
    Write-Log "Exiting script as Office is already installed."
    exit 0
}

Write-Log "Creating registry key to avoid download failing in restricted countries"
New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Force
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -Value "std::wstring|US" -PropertyType String -Force

# Download Office Deployment Tool setup.exe
$setupPath = "C:\OEM\setup.exe"
Write-Log "Checking for setup.exe"
if (Test-Path $setupPath) {
    Write-Log "setup.exe already exists, skipping download."
} else {
    Write-Log "Downloading Office Deployment Tool..."
    try {
        Invoke-WebRequest -Uri "http://go.microsoft.com/fwlink/?LinkID=829801" -OutFile $setupPath -ErrorAction Stop
        Write-Log "Downloaded setup.exe from primary URL."
    } catch {
        Write-Log "Failed to download from primary URL. Trying first fallback URL..."
        try {
            Invoke-WebRequest -Uri "https://archive.org/download/setup_20250603/setup.exe" -OutFile $setupPath -ErrorAction Stop
            Write-Log "Downloaded setup.exe from first fallback URL."
        } catch {
            Write-Log "Failed to download from first fallback URL. Trying second fallback URL..."
            try {
                $odtPath = "C:\OEM\odt.exe"
                Invoke-WebRequest -Uri "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_18730-20142.exe" -OutFile $odtPath -ErrorAction Stop
                Write-Log "Downloaded odt.exe from second fallback URL."
                
                # Extract setup.exe from odt.exe
                Write-Log "Extracting setup.exe from odt.exe..."
                $extractPath = "C:\OEM\ODT"
                New-Item -Path $extractPath -ItemType Directory -Force | Out-Null
                Start-Process -FilePath $odtPath -ArgumentList "/extract:$extractPath /quiet" -Wait -NoNewWindow
                if (Test-Path "$extractPath\setup.exe") {
                    Move-Item -Path "$extractPath\setup.exe" -Destination $setupPath -Force
                    Write-Log "Successfully extracted setup.exe."
                } else {
                    Write-Log "Failed to extract setup.exe from odt.exe."
                }
                # Clean up
                Remove-Item -Path $odtPath -Force -ErrorAction SilentlyContinue
                Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Failed to download or extract from second fallback URL."
            }
        }
    }
}

# Execute setup.exe with OfficeConfiguration.xml
Write-Log "Running Office installation..."
try {
    $process = Start-Process -FilePath $setupPath -ArgumentList "/configure C:\OEM\OfficeConfiguration.xml" -Wait -NoNewWindow -PassThru
    Write-Log "setup.exe process completed with exit code: $($process.ExitCode)"
} catch {
    Write-Log "Error running setup.exe: $_"
}

# Check for EXCEL.EXE to confirm installation
Write-Log "Checking for EXCEL.EXE..."
$installSuccess = $false
if (Test-Path "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") {
    Write-Log "EXCEL.EXE found. Office installation successful."
    $installSuccess = $true
} else {
    Write-Log "EXCEL.EXE not found. Office installation failed."
}

# Create success file
Write-Log "Creating success file..."
if ($installSuccess) {
    try {
        New-Item -Path "C:\OEM\success" -ItemType File -Force | Out-Null
        Write-Log "Success file created."
    } catch {
        Write-Log "Failed to create success file."
    }
} else {
    Write-Log "Skipping success file creation due to installation failure."
}

# Disable AutoLogon that was set up by install.bat
Write-Log "Disabling AutoLogon"
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name "AutoAdminLogon" -Value "0" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name "DefaultUserName" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name "DefaultPassword" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name "AutoLogonCount" -ErrorAction SilentlyContinue

Write-Log "Disabled AutoLogon in registry"
Write-Log "InstallOffice.ps1 completed, will reboot now"

# Initiate final reboot
Write-Log "Initiating final reboot"
Restart-Computer -Force

