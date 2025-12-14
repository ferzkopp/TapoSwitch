# TapoSwitch

**A Windows system tray application for controlling a single TP-Link Tapo smart device switch.**

TapoSwitch is a lightweight .NET 9 Windows application that provides quick access to your Tapo smart plugs and switches directly from the system tray. 
With a single click, you can toggle your devices on or off without opening the Tapo mobile app or web interface.

This is useful for turning external audio equipment (monitors, mixers) that are connected to a PC on and off.

## Features

- 🖱️ **One-Click Control**: Toggle your Tapo device on/off from the system tray
- 🚀 **Auto-Start**: Optional automatic startup at Windows login
- 🔄 **Auto-Shutdown**: Automatically turns off device when you log off or shut down Windows
- 🔒 **Secure**: Supports both SecurePassthrough and KLAP protocols
- 💾 **Lightweight**: Minimal resource usage, runs quietly in the background
- ⚙️ **Easy Installation**: Automated installer with interactive configuration
- 🔧 **Configurable**: Set device credentials and IP address during installation

## Quick Start

### Prerequisites

- Windows 10/11
- Visual Studio 2026
- .NET 9 Runtime
- Administrator privileges (for installation)
- TP-Link Tapo smart device (plug or switch)
- Tapo device registered and on known LAN IP address

### Installation Steps

#### 1. Build the Application

```bash
dotnet build -c Release
```

**Note**: The project automatically downloads the TapoConnect NuGet package (v3.2.4) and dependencies during the build.

This creates the executable at: `TapoSwitch\bin\Release\net9.0-windows\TapoSwitch.exe`

#### 2. Run the Installer

**Option A: Double-Click Install (Recommended)**
- Double-click `Install-Startup.bat`
- Choose "Install"
- Follow the prompts to configure your device

**Option B: PowerShell**
```powershell
# Run PowerShell as Administrator
.\Install-Startup.ps1
```

**The installer will:**
- Display current configuration settings
- Let you edit Tapo credentials (username, password, IP address)
- Optionally install to Program Files (default, recommended)
- Create Task Scheduler entry for automatic startup at login

## What Gets Installed

- **Task Scheduler** entry to start TapoSwitch at user login
- **Program Files** installation (optional, default: `C:\Program Files\TapoSwitch`)
- **Configuration** stored in `TapoSwitch.dll.config`

## Usage

Once installed and configured:

1. **System Tray Icon**: Look for the TapoSwitch icon in your system tray
   - 💡 Icon shows current device state (on/off)
2. **Left Click**: Toggle device on/off
3. **Right Click**: Access context menu (Exit option)
4. **Auto-Start**: Application starts automatically at login (if installed)

### Automatic Device Control

TapoSwitch automatically manages your device in the following scenarios:

- **Startup**: Device is turned OFF when the application starts (to ensure a known state)
- **Log Off**: Device is turned OFF when you log off Windows
- **Shutdown**: Device is turned OFF when you shut down or restart Windows
- **Exit**: Device is turned OFF when you exit the application from the tray menu

This ensures your Tapo device (e.g., a desk lamp, monitor light, or PC-connected equipment) is automatically controlled based on your computer usage patterns.

**Note**: The device will NOT automatically turn off when you lock your screen (Win+L). You can manually control the device at any time using the system tray icon.

## Configuration

### During Installation

The installer allows you to configure:
- **Username**: Your Tapo account email
- **Password**: Your Tapo account password
- **IP Address**: Local IP address of your Tapo device

### After Installation

Edit the configuration file at the appropriate location:

**If installed to Program Files:**
```
C:\Program Files\TapoSwitch\TapoSwitch.dll.config
```

**If using build folder:**
```
TapoSwitch\bin\Release\net9.0-windows\TapoSwitch.dll.config
```

**Configuration Format:**
```xml
<configuration>
  <appSettings>
    <add key="Username" value="your-tapo-email@example.com"/>
    <add key="Password" value="your-tapo-password"/>
    <add key="IpAddress" value="192.168.1.100"/>
    <add key="ShutdownTimeoutSeconds" value="2"/>
  </appSettings>
</configuration>
```

**Configuration Parameters:**

| Parameter | Description | Default | Valid Values |
|-----------|-------------|---------|--------------|
| `Username` | Your Tapo account email address | *(required)* | Any valid email |
| `Password` | Your Tapo account password | *(required)* | Any string |
| `IpAddress` | Local IP address of your Tapo device | *(required)* | Valid IPv4 address |
| `ShutdownTimeoutSeconds` | Timeout in seconds to wait for device to turn off during system shutdown/logoff | `2` | Positive integer (1-10 recommended) |

**About ShutdownTimeoutSeconds:**
- Controls how long the application waits for the device to turn off when Windows is shutting down or logging off
- A shorter timeout (1-2 seconds) is recommended for local network devices
- Increase the value (3-5 seconds) if you have a slower network or the device is not turning off reliably
- If the timeout is reached, Windows shutdown/logoff proceeds anyway to prevent hanging

**Note**: Restart the application after changing configuration.

## Installation Options

### Advanced Installation

**Install without copying to Program Files:**
```powershell
.\Install-Startup.ps1 -NoInstallToPrograms
```

**Install with custom startup delay:**
```powershell
# Delay startup by 15 seconds after login (default is 10 seconds)
.\Install-Startup.ps1 -StartupDelay 15
```

**Combine options:**
```powershell
# Install to build folder with 5 second startup delay
.\Install-Startup.ps1 -NoInstallToPrograms -StartupDelay 5
```

**Uninstall:**
```powershell
.\Install-Startup.ps1 -Uninstall
```
or double-click `Install-Startup.bat` and choose "Uninstall"

**Installation Parameters:**

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-AppName` | Name of the application (used for Task Scheduler and Program Files folder) | `TapoSwitch` |
| `-StartupDelay` | Delay in seconds before starting the app after login (allows system tray to initialize) | `10` |
| `-NoInstallToPrograms` | Skip copying files to Program Files, run from build folder instead | *(not set)* |
| `-Uninstall` | Remove the application and scheduled task | *(not set)* |

## Verification

After installation:

1. **Check if installed**:
   - Task Scheduler: Press `Win + R`, type `taskschd.msc`, look for "TapoSwitch" task
   - Program Files: Check `C:\Program Files\TapoSwitch` for installed files (if applicable)

2. **Test it**:
   - Log out and log back in
   - Check system tray for TapoSwitch icon
   - Or run manually: `C:\Program Files\TapoSwitch\TapoSwitch.exe`

## Troubleshooting

### Application Doesn't Start

1. **Check configuration**: Verify the config file has correct credentials and IP address:
   - **Program Files**: `C:\Program Files\TapoSwitch\TapoSwitch.dll.config`
   - **Build folder**: `TapoSwitch\bin\Release\net9.0-windows\TapoSwitch.dll.config`
2. **Check network**: Ensure device is reachable from your PC
3. **Check logs**: Look in Event Viewer → Windows Logs → Application
4. **Run manually**: Double-click `TapoSwitch.exe` to see if there are errors

### Configuration Changes Not Working

If you change configuration and it doesn't take effect:
- Make sure you edited the **correct config file** (see locations above)
- NOT the source file `TapoSwitch\App.config`
- Restart the application after changes
- You can re-run the installer to update configuration interactively

### Device Not Responding

- Verify the device IP address is correct
- Ensure the device is powered on and connected to your network
- Check that your Tapo credentials are valid
- Try pinging the device IP to verify network connectivity

### Device Not Turning Off During Shutdown

If your device doesn't turn off when shutting down or logging off:

1. **Increase the timeout**: Edit `TapoSwitch.dll.config` and increase `ShutdownTimeoutSeconds`:
   ```xml
   <add key="ShutdownTimeoutSeconds" value="5"/>
   ```
2. **Check network speed**: Slow network connections may need more time
3. **Verify device is reachable**: Test by manually toggling the device before shutdown
4. **Check Event Viewer**: Look for errors in Windows Event Viewer → Application logs
5. **Test the installer delay**: The Task Scheduler startup delay (default 10 seconds) may need adjustment:
   ```powershell
   .\Install-Startup.ps1 -StartupDelay 15
   ```

### Permission Issues

If you get "Access Denied" errors:
- Run PowerShell as Administrator
- The installation script requires admin rights for Task Scheduler and Program Files
- Use `-NoInstallToPrograms` flag if you cannot install to Program Files

### Wrong Executable Used

The installer looks for executables in this order:
1. `bin\Release\net9.0-windows\TapoSwitch.exe` (preferred)
2. `bin\Debug\net9.0-windows\TapoSwitch.exe` (fallback)

Always build in Release mode for production:
```bash
dotnet build -c Release
```

## Uninstalling

Run any of these:

```bash
# Batch file
Install-Startup.bat
# Then choose option 2 (Uninstall)

# PowerShell (as Admin)
.\Install-Startup.ps1 -Uninstall
```

The uninstall will:
- Stop any running TapoSwitch instances
- Remove the scheduled task
- Remove files from Program Files (if installed there)

## Security Notes

⚠️ **Important**: Your Tapo credentials are stored in plain text in the configuration file.

**Configuration file locations:**
- **Program Files**: `C:\Program Files\TapoSwitch\TapoSwitch.dll.config`
- **Build folder**: `TapoSwitch\bin\Release\net9.0-windows\TapoSwitch.dll.config`

**Security recommendations:**
1. **Never commit the built output folder** (`bin\`, `obj\`) to version control
2. Set appropriate file permissions on the config file to restrict access
3. If installed to Program Files, only administrators can modify the config by default
4. Keep your build output folder excluded from backups or cloud sync services
5. Consider using Windows Credential Manager in future versions
6. The source `TapoSwitch\App.config` file should remain with placeholder values only

**Best Practice:**
- Use the interactive installer to configure credentials securely
- Keep `TapoSwitch\App.config` with placeholder values (safe for repository)
- Create a `.gitignore` entry for `bin/` and `obj/` folders (should already be present)
- Install to Program Files for better security (admin rights required to modify)

### For Developers

When working on the codebase, you can create an `App.config.user` file to store your credentials without risk of accidentally committing them:

1. **Copy the template** (in the root directory alongside `Install-Startup.ps1`):
   ```bash
   copy App.config App.config.user
   ```

2. **Edit `App.config.user`** with your actual credentials:
   ```xml
   <configuration>
     <appSettings>
       <add key="Username" value="your-email@example.com"/>
       <add key="Password" value="your-actual-password"/>
       <add key="IpAddress" value="192.168.1.100"/>
       <add key="ShutdownTimeoutSeconds" value="2"/>
     </appSettings>
   </configuration>
   ```

3. **How it works**:
   - Place `App.config.user` in the root directory (same location as `Install-Startup.ps1`)
   - When you run the installer, it automatically detects `App.config.user`
   - The installer copies `App.config.user` to the build output folder as `TapoSwitch.dll.config`
   - Then uses that configuration for installation
   - `App.config.user` is in `.gitignore` so it won't be committed
   - When you build the project, your build uses the standard `App.config` (with placeholders)

4. **Keep `App.config` with placeholders**:
   ```xml
   <add key="Username" value="[YourTapoRegistrationEmail]"/>
   <add key="Password" value="[YourTapoRegistrationPassword]"/>
   <add key="IpAddress" value="[YourTapoSwitchIPOnTheLAN]"/>
   ```

This approach allows you to:
- Work with real credentials locally without modifying `App.config`
- Install and test the application during development without re-entering credentials
- Prevent accidentally committing credentials to version control
- Keep the repository safe for sharing

**Note**: When the installer detects `App.config.user`, it will show "[Dev Mode]" and automatically use those credentials for the installation.

**File Location**: Make sure `App.config.user` is in the root project directory:
```
TapoSwitch/
├── App.config              # Template with placeholders (tracked in git)
├── App.config.user         # Your credentials (ignored by git)
├── Install-Startup.ps1
├── Install-Startup.bat
└── ...
```

## Advanced Configuration

### Change Task Scheduler Settings

After installation, you can modify the task in Task Scheduler:
- Change trigger (e.g., add delay after logon)
- Change conditions (e.g., only on AC power)
- Configure for multiple users

### Running as a Different User

To run TapoSwitch as a different user:
1. Open Task Scheduler
2. Find the TapoSwitch task
3. Edit the task
4. Change the user account in "General" tab

### Finding Your Device IP Address

To find your Tapo device's IP address:
1. Open the Tapo mobile app
2. Go to device settings
3. Look for "Device Info" or similar
4. Note the IP address displayed

Or check your router's connected devices list.

## Technical Details

- **Framework**: .NET 9 (Windows)
- **UI Framework**: Windows Forms
- **Protocols Supported**: SecurePassthrough, KLAP
- **Target OS**: Windows 10/11
- **Dependencies**: 
  - [TapoConnect](https://www.nuget.org/packages/TapoConnect/) v3.2.4 (NuGet package)
  - Microsoft.CSharp v4.7.0
  - System.Data.DataSetExtensions v4.5.0

## Project Structure

```
TapoSwitch/
├── Program.cs              # Main application entry point
├── App.config              # Source configuration (placeholders)
├── Install-Startup.ps1     # PowerShell installation script
├── Install-Startup.bat     # Batch installation wrapper
├── README.md               # This file
└── bin/Release/            # Build output folder
    └── net9.0-windows/
        ├── TapoSwitch.exe
        └── TapoSwitch.dll.config  # Runtime configuration
```

## Support

If you encounter issues:
1. Check the Troubleshooting section above
2. Verify your build is up to date
3. Check the [GitHub repository](https://github.com/cwakefie27/TapoConnect) for issues
4. Create a new issue with logs and error messages

## Contributing

This project uses the [TapoConnect NuGet package](https://www.nuget.org/packages/TapoConnect/) (v3.2.4) for device communication.

For more information about TapoConnect, visit the [GitHub repository](https://github.com/cwakefie27/TapoConnect).

## License

See the LICENSE file in the repository for details.

---

**Made with ❤️ for TP-Link Tapo device automation**
