@echo off
REM TapoSwitch Startup Installer (Batch Script)
REM This provides a simple double-click installation option

REM Check for administrator privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator privileges...
    powershell.exe -ExecutionPolicy Bypass -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo ================================================
echo   TapoSwitch Installation
echo ================================================
echo.

REM Check for executable
set "RELEASE_EXE=%~dp0bin\Release\net9.0-windows\TapoSwitch.exe"
set "DEBUG_EXE=%~dp0bin\Debug\net9.0-windows\TapoSwitch.exe"

if exist "%RELEASE_EXE%" (
    set "EXE_PATH=%RELEASE_EXE%"
    echo Found Release build
) else if exist "%DEBUG_EXE%" (
    set "EXE_PATH=%DEBUG_EXE%"
    echo Found Debug build
) else (
    echo ERROR: TapoSwitch.exe not found!
    echo Please build the project first.
    echo.
    echo Build command: dotnet build -c Release
    echo.
    pause
    exit /b 1
)

echo Executable: %EXE_PATH%
echo.

REM Offer choice
echo Choose an option:
echo   1. Install (with configuration)
echo   2. Uninstall
echo   3. Cancel
echo.
choice /C 123 /N /M "Select option (1-3): "

if errorlevel 3 goto :cancel
if errorlevel 2 goto :uninstall
if errorlevel 1 goto :install

:install
echo.
echo Installing TapoSwitch (requires admin)...
powershell.exe -NoExit -ExecutionPolicy Bypass -File "%~dp0Install-Startup.ps1"
goto :done

:uninstall
echo.
echo Uninstalling TapoSwitch (requires admin)...
powershell.exe -NoExit -ExecutionPolicy Bypass -File "%~dp0Install-Startup.ps1" -Uninstall
goto :done

:cancel
echo.
echo Installation cancelled.
goto :end

:done
echo.
echo ================================================
echo   Done
echo ================================================
echo.

:end
pause

