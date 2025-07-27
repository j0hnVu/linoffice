@echo off
setlocal EnableDelayedExpansion

:: Check if Ethernet adapter exists and is connected
netsh interface show interface name="Ethernet" | findstr /C:"Connected" >nul
if %ERRORLEVEL% neq 0 (
    echo Error: Adapter "Ethernet" not found or not connected. Check adapter status with "netsh interface show interface".
    pause
    exit /b
)

echo Found adapter: Ethernet

:: Set DNS to DHCP
netsh interface ipv4 set dns name="Ethernet" source=dhcp
if %ERRORLEVEL% neq 0 (
    echo Error: Failed to set DNS to DHCP. Ensure script is run as Administrator.
) else (
    echo DNS set to Automatic (DHCP) for Ethernet adapter.
)

pause
endlocal