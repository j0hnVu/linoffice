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

:: Set DNS to 127.0.0.1
netsh interface ipv4 set dns name="Ethernet" source=static addr=127.0.0.1
if %ERRORLEVEL% neq 0 (
    echo Error: Failed to set DNS to 127.0.0.1. Ensure script is run as Administrator.
) else (
    echo DNS set to Manual (127.0.0.1) for Ethernet adapter.
)

:: Flush DNS cache
ipconfig /flushdns
if %ERRORLEVEL% neq 0 (
    echo Warning: Failed to flush DNS cache. Ensure script is run as Administrator.
) else (
    echo DNS cache flushed successfully.
)

pause
endlocal