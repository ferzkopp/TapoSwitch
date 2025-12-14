# TapoSwitch Startup Installation Script
# This script installs TapoSwitch to run automatically at Windows startup using Task Scheduler

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$AppName = 'TapoSwitch',
    
    [Parameter(Mandatory=$false)]
    [switch]$Uninstall,
    
    [Parameter(Mandatory=$false)]
    [switch]$NoInstallToPrograms,
    
    [Parameter(Mandatory=$false)]
    [int]$StartupDelay = 10
)

# Get the directory where this script is located
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Determine the executable path
$BuildPath = Join-Path $ScriptDir "bin\Release\net9.0-windows"
$DebugPath = Join-Path $ScriptDir "bin\Debug\net9.0-windows"
$BuildExe = Join-Path $BuildPath "TapoSwitch.exe"
$DebugExe = Join-Path $DebugPath "TapoSwitch.exe"

$SourcePath = $null
if (Test-Path $BuildExe) {
    $SourcePath = $BuildPath
    Write-Host "Found Release build: $BuildPath" -ForegroundColor Green
} elseif (Test-Path $DebugExe) {
    $SourcePath = $DebugPath
    Write-Host "Found Debug build: $DebugPath" -ForegroundColor Yellow
    Write-Warning "Using Debug build. Consider building in Release mode for production use."
} else {
    Write-Error "TapoSwitch.exe not found. Please build the project first."
    Write-Host "Expected locations:" -ForegroundColor Yellow
    Write-Host "  - $BuildExe" -ForegroundColor Yellow
    Write-Host "  - $DebugExe" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To build the project, run:" -ForegroundColor Cyan
    Write-Host "  dotnet build -c Release" -ForegroundColor White
    exit 1
}

# Function to stop running TapoSwitch instances
function Stop-TapoSwitchProcesses {
    Write-Host ""
    Write-Host "Checking for running TapoSwitch instances..." -ForegroundColor Cyan
    
    $RunningProcesses = Get-Process -Name "TapoSwitch" -ErrorAction SilentlyContinue
    if ($RunningProcesses) {
        Write-Host "Found $($RunningProcesses.Count) running instance(s). Stopping..." -ForegroundColor Yellow
        $RunningProcesses | Stop-Process -Force
        Start-Sleep -Seconds 1
        Write-Host "[OK] Stopped all TapoSwitch instances" -ForegroundColor Green
    } else {
        Write-Host "No running instances found" -ForegroundColor Gray
    }
}

# Function to read configuration
function Get-ConfigValue {
    param($ConfigPath, $Key)
    
    if (-not (Test-Path $ConfigPath)) {
        return $null
    }
    
    [xml]$config = Get-Content $ConfigPath
    $setting = $config.configuration.appSettings.add | Where-Object { $_.key -eq $Key }
    return $setting.value
}

# Function to set configuration value
function Set-ConfigValue {
    param($ConfigPath, $Key, $Value)
    
    [xml]$config = Get-Content $ConfigPath
    $setting = $config.configuration.appSettings.add | Where-Object { $_.key -eq $Key }
    
    if ($setting) {
        $setting.value = $Value
    } else {
        $newSetting = $config.CreateElement("add")
        $newSetting.SetAttribute("key", $Key)
        $newSetting.SetAttribute("value", $Value)
        $config.configuration.appSettings.AppendChild($newSetting) | Out-Null
    }
    
    $config.Save($ConfigPath)
}


# Function to configure settings interactively
function Edit-Configuration {
    param($ConfigPath)
    
    Write-Host ""
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  Current Configuration" -ForegroundColor Cyan
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    
    $username = Get-ConfigValue -ConfigPath $ConfigPath -Key "Username"
    $password = Get-ConfigValue -ConfigPath $ConfigPath -Key "Password"
    $ipAddress = Get-ConfigValue -ConfigPath $ConfigPath -Key "IpAddress"
    
    Write-Host "Username  : $username" -ForegroundColor White
    Write-Host "Password  : $password" -ForegroundColor White
    Write-Host "IP Address: $ipAddress" -ForegroundColor White
    Write-Host ""
    
    $response = Read-Host "Do you want to modify these settings? (y/N)"
    
    if ($response -eq 'y' -or $response -eq 'Y') {
        Write-Host ""
        Write-Host "Enter new values (press Enter to keep current value):" -ForegroundColor Cyan
        Write-Host ""
        
        $newUsername = Read-Host "Tapo Account Username [$username]"
        if ([string]::IsNullOrWhiteSpace($newUsername)) {
            $newUsername = $username
        }
        
        $newPassword = Read-Host "Tapo Account Password [$(if($password){'***'}else{''})]"
        if ([string]::IsNullOrWhiteSpace($newPassword)) {
            $newPassword = $password
        }
        
        $newIpAddress = Read-Host "Device IP Address [$ipAddress]"
        if ([string]::IsNullOrWhiteSpace($newIpAddress)) {
            $newIpAddress = $ipAddress
        }
        
        Write-Host ""
        Write-Host "Updating configuration..." -ForegroundColor Cyan
        
        Set-ConfigValue -ConfigPath $ConfigPath -Key "Username" -Value $newUsername
        Set-ConfigValue -ConfigPath $ConfigPath -Key "Password" -Value $newPassword
        Set-ConfigValue -ConfigPath $ConfigPath -Key "IpAddress" -Value $newIpAddress
        
        Write-Host "[OK] Configuration updated" -ForegroundColor Green
        
        return $true
    }
    
    return $false
}

# Function to copy files to Program Files
function Install-ToPrograms {
    param($SourcePath, $AppName)
    
    $ProgramFiles = [Environment]::GetFolderPath("ProgramFiles")
    $DestPath = Join-Path $ProgramFiles $AppName
    
    Write-Host ""
    Write-Host "Installing to: $DestPath" -ForegroundColor Cyan
    
    # Create directory if it doesn't exist
    if (-not (Test-Path $DestPath)) {
        New-Item -ItemType Directory -Path $DestPath -Force | Out-Null
    }
    
    # Copy all files
    Write-Host "Copying files..." -ForegroundColor Cyan
    Copy-Item -Path "$SourcePath\*" -Destination $DestPath -Recurse -Force
    
    Write-Host "[OK] Files copied to Program Files" -ForegroundColor Green
    
    return $DestPath
}

# Function to uninstall from Program Files
function Uninstall-FromPrograms {
    param($AppName)
    
    $ProgramFiles = [Environment]::GetFolderPath("ProgramFiles")
    $InstallPath = Join-Path $ProgramFiles $AppName
    
    if (Test-Path $InstallPath) {
        Write-Host "Removing from Program Files..." -ForegroundColor Cyan
        Remove-Item -Path $InstallPath -Recurse -Force
        Write-Host "[OK] Removed from Program Files" -ForegroundColor Green
    } else {
        Write-Host "! Not found in Program Files" -ForegroundColor Yellow
    }
}

# Function to enable Task Scheduler history
function Enable-TaskSchedulerHistory {
    Write-Host "Enabling Task Scheduler history..." -ForegroundColor Cyan
    
    try {
        # Enable history via registry or WEvtUtil
        $result = & wevtutil.exe set-log Microsoft-Windows-TaskScheduler/Operational /enabled:true 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Task Scheduler history enabled" -ForegroundColor Green
        } else {
            Write-Warning "Could not enable Task Scheduler history. You may need to enable it manually in Task Scheduler."
        }
    } catch {
        Write-Warning "Could not enable Task Scheduler history: $_"
    }
}

# Function to install using Task Scheduler
function Install-TaskScheduler {
    param(
        $ExePath, 
        $AppName,
        $StartupDelay
    )
    
    Write-Host ""
    Write-Host "Creating scheduled task..." -ForegroundColor Cyan
    
    $TaskName = $AppName
    $TaskDescription = "Automatically starts TapoSwitch to control Tapo smart devices"
    
    # Check if task already exists
    $ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($ExistingTask) {
        Write-Host "! Task already exists. Removing old task..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }
    
    # Create the action
    $Action = New-ScheduledTaskAction -Execute $ExePath -WorkingDirectory (Split-Path -Parent $ExePath)
    
    # Create the trigger (at logon) with delay to allow system tray to initialize
    $Trigger = New-ScheduledTaskTrigger -AtLogOn
    
    # Add delay to the trigger to allow Windows Explorer/System Tray to fully initialize
    if ($StartupDelay -gt 0) {
        $Trigger.Delay = "PT$($StartupDelay)S"
        Write-Host "Startup delay configured: $StartupDelay seconds" -ForegroundColor Gray
    }
    
    # Create the principal (run as current user with Interactive logon for desktop/tray icon access)
    $Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
    
    # Create the settings
    $Settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -DontStopOnIdleEnd `
        -ExecutionTimeLimit 0 `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 1)
    
    # Register the task
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Description $TaskDescription `
        -Action $Action `
        -Trigger $Trigger `
        -Principal $Principal `
        -Settings $Settings | Out-Null
    
    Write-Host "[OK] Scheduled task created: $TaskName" -ForegroundColor Green
    if ($StartupDelay -gt 0) {
        Write-Host "[OK] $AppName will start $StartupDelay seconds after logon" -ForegroundColor Green
    } else {
        Write-Host "[OK] $AppName will start immediately at logon" -ForegroundColor Green
    }
    
    # Enable history for Task Scheduler
    Enable-TaskSchedulerHistory
}

# Function to run the scheduled task
function Start-ScheduledTaskNow {
    param($AppName)
    
    Write-Host ""
    Write-Host "Starting scheduled task..." -ForegroundColor Cyan
    
    $TaskName = $AppName
    
    try {
        Start-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        Write-Host "[OK] Task started successfully" -ForegroundColor Green
        Start-Sleep -Seconds 2
        
        # Check if process is running
        $Process = Get-Process -Name "TapoSwitch" -ErrorAction SilentlyContinue
        if ($Process) {
            Write-Host "[OK] TapoSwitch is now running (PID: $($Process.Id))" -ForegroundColor Green
        } else {
            Write-Warning "Task started but TapoSwitch process not detected. Check Task Scheduler history for errors."
        }
        
        # Show last run result
        $Task = Get-ScheduledTask -TaskName $TaskName
        $TaskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
        Write-Host ""
        Write-Host "Last Run Time: $($TaskInfo.LastRunTime)" -ForegroundColor Gray
        Write-Host "Last Result  : 0x$($TaskInfo.LastTaskResult.ToString('X'))" -ForegroundColor Gray
        Write-Host "Next Run Time: $($TaskInfo.NextRunTime)" -ForegroundColor Gray
        
    } catch {
        Write-Error "Failed to start task: $_"
    }
}

# Function to uninstall from Task Scheduler
function Uninstall-TaskScheduler {
    param($AppName)
    
    Write-Host "Removing scheduled task..." -ForegroundColor Cyan
    
    $TaskName = $AppName
    
    try {
        # First, list all tasks with similar names for debugging
        Write-Host "Searching for tasks matching '$TaskName'..." -ForegroundColor Gray
        $AllTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "*$TaskName*" }
        if ($AllTasks) {
            Write-Host "Found $($AllTasks.Count) matching task(s):" -ForegroundColor Gray
            foreach ($t in $AllTasks) {
                Write-Host "  - $($t.TaskPath)$($t.TaskName) (State: $($t.State))" -ForegroundColor Gray
            }
        }
        
        # Try to get the exact task
        $Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        
        if ($Task) {
            Write-Host "Found task: $($Task.TaskPath)$($Task.TaskName)" -ForegroundColor Gray
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
            Write-Host "[OK] Removed scheduled task: $TaskName" -ForegroundColor Green
            
            # Verify removal
            Start-Sleep -Milliseconds 500
            $VerifyTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
            if ($VerifyTask) {
                Write-Warning "Task still exists after removal attempt. Trying alternative method..."
                
                # Try using schtasks.exe as fallback
                $result = & schtasks.exe /Delete /TN $TaskName /F 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "[OK] Removed task using schtasks.exe" -ForegroundColor Green
                } else {
                    Write-Error "Failed to remove task: $result"
                }
            }
        } else {
            Write-Host "! Scheduled task '$TaskName' not found" -ForegroundColor Yellow
            
            # Double-check with schtasks.exe
            Write-Host "Verifying with schtasks.exe..." -ForegroundColor Gray
            $result = & schtasks.exe /Query /TN $TaskName 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Warning "Task found via schtasks.exe. Attempting removal..."
                $deleteResult = & schtasks.exe /Delete /TN $TaskName /F 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "[OK] Removed task using schtasks.exe" -ForegroundColor Green
                } else {
                    Write-Error "Failed to remove task: $deleteResult"
                }
            } else {
                # Check if any of the similar tasks should be removed
                if ($AllTasks) {
                    Write-Host ""
                    Write-Host "Found similar tasks. Would you like to remove them? (y/N)" -ForegroundColor Yellow
                    foreach ($t in $AllTasks) {
                        Write-Host "  - $($t.TaskPath)$($t.TaskName)" -ForegroundColor Yellow
                    }
                    # For unattended operation, we won't prompt - just list them
                    Write-Host "If you need to remove these tasks manually, use:" -ForegroundColor Yellow
                    Write-Host "  Unregister-ScheduledTask -TaskName '<TaskName>' -Confirm:`$false" -ForegroundColor White
                }
            }
        }
    } catch {
        Write-Error "Error during task removal: $_"
        
        # Try fallback method
        Write-Host "Attempting removal with schtasks.exe..." -ForegroundColor Yellow
        $result = & schtasks.exe /Delete /TN $TaskName /F 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Removed task using schtasks.exe" -ForegroundColor Green
        } else {
            Write-Error "Failed to remove task with fallback method: $result"
        }
    }
}


# Main execution
Write-Host ""
Write-Host "---------------------------------------------------" -ForegroundColor Cyan
Write-Host "  TapoSwitch Installation" -ForegroundColor Cyan
Write-Host "---------------------------------------------------" -ForegroundColor Cyan
Write-Host ""

if ($Uninstall) {
    Write-Host "Mode: Uninstall" -ForegroundColor Yellow
    
    # Stop running instances
    Stop-TapoSwitchProcesses
    
    Write-Host ""
    
    # Uninstall scheduled task
    Uninstall-TaskScheduler -AppName $AppName
    
    # Uninstall from Program Files
    Uninstall-FromPrograms -AppName $AppName
    
    Write-Host ""
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  Uninstallation Complete!" -ForegroundColor Green
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    
} else {
    Write-Host "Mode: Install" -ForegroundColor Green
    Write-Host "Startup Delay: $StartupDelay seconds" -ForegroundColor Gray
    
    # Stop running instances before installation
    Stop-TapoSwitchProcesses
    
    Write-Host ""
    
    # Check for App.config.user in the script directory (for development)
    $UserConfigPath = Join-Path $ScriptDir "App.config.user"
    
    if (Test-Path $UserConfigPath) {
        Write-Host ""
        Write-Host "[Dev Mode] Found App.config.user in script directory" -ForegroundColor Magenta
        Write-Host "Source: $UserConfigPath" -ForegroundColor Gray
        
        # Determine the target config file path in the build output
        $TargetConfigPath = Join-Path $SourcePath "TapoSwitch.dll.config"
        
        # Copy App.config.user to the build output as TapoSwitch.dll.config
        Copy-Item -Path $UserConfigPath -Destination $TargetConfigPath -Force
        Write-Host "[OK] Copied App.config.user to build output" -ForegroundColor Green
        Write-Host "Target: $TargetConfigPath" -ForegroundColor Gray
    }
    
    # Determine config file path
    $ConfigPath = Join-Path $SourcePath "TapoSwitch.dll.config"
    
    if (-not (Test-Path $ConfigPath)) {
        Write-Error "Configuration file not found: $ConfigPath"
        exit 1
    }
    
    # Show and optionally edit configuration
    Edit-Configuration -ConfigPath $ConfigPath
    
    # Determine installation path
    $InstallToPrograms = -not $NoInstallToPrograms
    $ExePath = $null
    $InstallPath = $null
    
    if ($InstallToPrograms) {
        Write-Host ""
        Write-Host "---------------------------------------------------" -ForegroundColor Cyan
        Write-Host "  Installation Location" -ForegroundColor Cyan
        Write-Host "---------------------------------------------------" -ForegroundColor Cyan
        
        $ProgramFiles = [Environment]::GetFolderPath("ProgramFiles")
        $DefaultPath = Join-Path $ProgramFiles $AppName
        
        Write-Host ""
        Write-Host "Default: Install to Program Files" -ForegroundColor White
        Write-Host "  Path: $DefaultPath" -ForegroundColor Gray
        Write-Host ""
        
        $response = Read-Host "Install to Program Files? (Y/n)"
        
        if ($response -eq 'n' -or $response -eq 'N') {
            $InstallToPrograms = $false
        }
    }
    
    if ($InstallToPrograms) {
        $InstallPath = Install-ToPrograms -SourcePath $SourcePath -AppName $AppName
        $ExePath = Join-Path $InstallPath "TapoSwitch.exe"
    } else {
        $ExePath = Join-Path $SourcePath "TapoSwitch.exe"
        $InstallPath = $SourcePath
        Write-Host ""
        Write-Host "Using build folder: $SourcePath" -ForegroundColor Cyan
    }
    
    # Install to Task Scheduler with startup delay
    Install-TaskScheduler -ExePath $ExePath -AppName $AppName -StartupDelay $StartupDelay
    
    Write-Host ""
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  Installation Complete!" -ForegroundColor Green
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Installation Details:" -ForegroundColor White
    Write-Host "  Location: $InstallPath" -ForegroundColor Gray
    Write-Host "  Executable: $ExePath" -ForegroundColor Gray
    Write-Host "  Config: $(Join-Path $InstallPath 'TapoSwitch.dll.config')" -ForegroundColor Gray
    Write-Host "  Startup Delay: $StartupDelay seconds" -ForegroundColor Gray
    Write-Host ""
    
    # Offer to run the task now
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    $response = Read-Host "Do you want to run the task now to test it? (Y/n)"
    
    if ($response -ne 'n' -and $response -ne 'N') {
        Start-ScheduledTaskNow -AppName $AppName
    }
    
    Write-Host ""
    Write-Host "Additional Options:" -ForegroundColor White
    Write-Host "  - Log off and log back in to test automatic startup" -ForegroundColor White
    Write-Host "  - The app will start $StartupDelay seconds after logon" -ForegroundColor White
    Write-Host "  - Run 'Start-ScheduledTask -TaskName $AppName' to start the task" -ForegroundColor White
    Write-Host "  - View the task in Task Scheduler (taskschd.msc)" -ForegroundColor White
    Write-Host "  - Check Task History in Task Scheduler for troubleshooting" -ForegroundColor White
    Write-Host ""
    Write-Host "To change startup delay, reinstall with:" -ForegroundColor White
    Write-Host "  .\Install-Startup.ps1 -StartupDelay <seconds>" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To uninstall, run:" -ForegroundColor White
    Write-Host "  .\Install-Startup.ps1 -Uninstall" -ForegroundColor Yellow
}

Write-Host ""

# Pause at the end if running from batch file or if window will close
if ($Host.Name -eq "ConsoleHost") {
    Write-Host "Press any key to close this window..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
