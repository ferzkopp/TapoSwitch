using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using TapoSwitch.Properties;
using System.Configuration;
using System.Net.Http;
using System.Net.Sockets;
using Microsoft.Win32;
using System.Threading;

// Uses Unofficial TP-Link Tapo smart device library for C#: https://github.com/cwakefie27/TapoConnect
using TapoConnect;
using TapoConnect.Protocol;
using TapoConnect.Dto;

namespace TapoSwitch
{
    /// <summary>
    /// Contains the main entry point and application context for the TapoSwitch application.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Gets the username from application settings.
        /// </summary>
        public static readonly string Username = ConfigurationManager.AppSettings["Username"];

        /// <summary>
        /// Gets the password from application settings.
        /// </summary>
        public static readonly string Password = ConfigurationManager.AppSettings["Password"];

        /// <summary>
        /// Gets the IP address of the Tapo device from application settings.
        /// </summary>
        public static readonly string IpAddress = ConfigurationManager.AppSettings["IpAddress"];

        /// <summary>
        /// Gets the shutdown timeout in seconds from application settings. Default is 2 seconds.
        /// </summary>
        public static readonly int ShutdownTimeoutSeconds = 
            int.TryParse(ConfigurationManager.AppSettings["ShutdownTimeoutSeconds"], out var timeout) && timeout > 0
                ? timeout
                : 2;

        /// <summary>
        /// Gets the connection retry attempts from application settings. Default is 3.
        /// </summary>
        public static readonly int ConnectionRetryAttempts = 
            int.TryParse(ConfigurationManager.AppSettings["ConnectionRetryAttempts"], out var retries) && retries > 0
                ? retries
                : 3;

        /// <summary>
        /// Gets the retry delay in milliseconds from application settings. Default is 2000ms.
        /// </summary>
        public static readonly int RetryDelayMilliseconds = 
            int.TryParse(ConfigurationManager.AppSettings["RetryDelayMilliseconds"], out var delay) && delay > 0
                ? delay
                : 2000;

        /// <summary>
        /// Maximum delay in milliseconds for exponential backoff. Default is 30000ms (30 seconds).
        /// </summary>
        private const int MaxRetryDelayMilliseconds = 30000;

        /// <summary>
        /// Maximum length for Windows tooltip text (63 characters + null terminator).
        /// </summary>
        private const int MaxTooltipLength = 63;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Validate configuration at startup
            if (!ValidateConfiguration(out string validationError))
            {
                MessageBox.Show(
                    $"Configuration Error: {validationError}\n\nPlease check your App.config file.",
                    "TapoSwitch Configuration Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Set the default font for the application (example: Microsoft Sans Serif, 8.25pt)
            using var font = new Font("Microsoft Sans Serif", 8.25F);
            Application.SetDefaultFont(font);

            var context = new CustomApplicationContext();
            Application.Run(context);
        }

        /// <summary>
        /// Validates critical configuration settings at startup.
        /// </summary>
        /// <param name="errorMessage">Error message if validation fails.</param>
        /// <returns>True if configuration is valid, false otherwise.</returns>
        private static bool ValidateConfiguration(out string errorMessage)
        {
            if (string.IsNullOrWhiteSpace(Username))
            {
                errorMessage = "Username is missing or empty.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(Password))
            {
                errorMessage = "Password is missing or empty.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(IpAddress))
            {
                errorMessage = "IpAddress is missing or empty.";
                return false;
            }

            // Validate IP address format
            if (!System.Net.IPAddress.TryParse(IpAddress, out _))
            {
                errorMessage = $"IpAddress '{IpAddress}' is not a valid IP address.";
                return false;
            }

            errorMessage = null;
            return true;
        }

        /// <summary>
        /// Provides the application context for the TapoSwitch tray application, including tray icon and device control.
        /// </summary>
        public class CustomApplicationContext : ApplicationContext, IDisposable
        {
            /// <summary>
            /// The tray icon displayed in the system tray.
            /// </summary>
            private NotifyIcon trayIcon;

            /// <summary>
            /// The client used to communicate with the Tapo device.
            /// </summary>
            private ITapoDeviceClient deviceClient;

            /// <summary>
            /// The authentication key for the Tapo device.
            /// </summary>
            private TapoDeviceKey deviceKey = null!;

            /// <summary>
            /// The information about the Tapo device.
            /// </summary>
            private DeviceGetInfoResult deviceInfo = null!;

            /// <summary>
            /// Indicates the current power state of the device.
            /// </summary>
            private volatile bool state = false;

            /// <summary>
            /// Indicates whether the application is shutting down.
            /// </summary>
            private volatile bool isShuttingDown = false;

            /// <summary>
            /// Indicates whether the device is currently connected.
            /// </summary>
            private volatile bool isConnected = false;

            /// <summary>
            /// Lock object for thread-safe operations.
            /// </summary>
            private readonly SemaphoreSlim connectionLock = new SemaphoreSlim(1, 1);

            /// <summary>
            /// Background timer for periodic connection checks.
            /// </summary>
            private System.Threading.Timer connectionCheckTimer;

            /// <summary>
            /// Synchronization context for UI thread marshaling.
            /// </summary>
            private readonly SynchronizationContext syncContext;

            /// <summary>
            /// Tracks whether Dispose has been called.
            /// </summary>
            private bool disposed = false;

            /// <summary>
            /// Initializes a new instance of the <see cref="CustomApplicationContext"/> class.
            /// Sets up the tray icon and starts device initialization.
            /// </summary>
            public CustomApplicationContext()
            {
                // Capture UI synchronization context
                syncContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();

                // Initialize Tray Icon
                ContextMenuStrip contextMenuStrip = new ContextMenuStrip();
                contextMenuStrip.Items.Add("Exit", null, Exit);

                string tooltipText = "Connecting to device...";
                trayIcon = new NotifyIcon()
                {
                    Icon = Resources.SwitchOff,
                    Text = tooltipText,
                    ContextMenuStrip = contextMenuStrip,
                    Visible = true
                };

                // Handle clicks
                trayIcon.MouseClick += MouseClick;

                // Handle application exit
                Application.ApplicationExit += OnApplicationExit;

                // Register for session change events (logoff, shutdown)
                SystemEvents.SessionEnding += OnSessionEnding;
                SystemEvents.SessionSwitch += OnSessionSwitch;

                // Start async initialization after message loop starts
                Task.Run(() => InitializeAsync());

                // Start periodic connection check (every 30 seconds)
                var connectionCheckInterval = TimeSpan.FromSeconds(30);
                connectionCheckTimer = new System.Threading.Timer(
                    ConnectionCheckCallback,
                    null,
                    connectionCheckInterval,
                    connectionCheckInterval);
            }

            /// <summary>
            /// Timer callback that safely handles connection checks.
            /// </summary>
            private async void ConnectionCheckCallback(object state)
            {
                try
                {
                    await EnsureConnectionAsync().ConfigureAwait(false);
                }
                catch
                {
                    // Prevent unhandled exceptions from crashing the app
                }
            }

            /// <summary>
            /// Asynchronously initializes the device connection and ensures the device is off at startup.
            /// Retries on failure until successful or application shutdown.
            /// </summary>
            private async Task InitializeAsync()
            {
                int attempt = 0;
                while (!isShuttingDown && !isConnected)
                {
                    attempt++;
                    try
                    {
                        await InitializeConnection().ConfigureAwait(false);
                        await SyncDeviceStateAsync().ConfigureAwait(false);
                        await TurnOffAsync().ConfigureAwait(false);
                        
                        isConnected = true;
                        InvokeOnUIThread(() =>
                        {
                            if (!isShuttingDown)
                            {
                                UpdateTrayIcon();
                                UpdateTooltip();
                            }
                        });
                        break;
                    }
                    catch (Exception ex) when (IsNetworkException(ex))
                    {
                        // Network issue - will retry
                        int delayMs;
                        if (attempt < ConnectionRetryAttempts)
                        {
                            // Exponential backoff capped at MaxRetryDelayMilliseconds
                            delayMs = Math.Min(RetryDelayMilliseconds * attempt, MaxRetryDelayMilliseconds);
                        }
                        else
                        {
                            // Keep retrying but with capped delay
                            delayMs = MaxRetryDelayMilliseconds;
                        }
                        await Task.Delay(delayMs).ConfigureAwait(false);
                    }
                    catch
                    {
                        // Non-network exception - wait and retry with capped delay
                        await Task.Delay(Math.Min(RetryDelayMilliseconds * 3, MaxRetryDelayMilliseconds)).ConfigureAwait(false);
                    }
                }
            }

            /// <summary>
            /// Asynchronously establishes a connection to the Tapo device and retrieves device information.
            /// </summary>
            public async Task InitializeConnection()
            {
                await connectionLock.WaitAsync().ConfigureAwait(false);
                try
                {
                    deviceClient = new TapoDeviceClient(new List<ITapoDeviceClient>
                    {
                        new SecurePassthroughDeviceClient(),
                        new KlapDeviceClient(),
                    });

                    deviceKey = await deviceClient.LoginByIpAsync(IpAddress, Username, Password).ConfigureAwait(false);
                    deviceInfo = await deviceClient.GetDeviceInfoAsync(deviceKey).ConfigureAwait(false);
                }
                finally
                {
                    connectionLock.Release();
                }
            }

            /// <summary>
            /// Ensures the connection is active, reconnecting if necessary.
            /// </summary>
            private async Task EnsureConnectionAsync()
            {
                if (isShuttingDown)
                    return;

                if (!isConnected)
                {
                    await InitializeAsync().ConfigureAwait(false);
                }
            }

            /// <summary>
            /// Synchronizes local state with actual device state.
            /// </summary>
            private async Task SyncDeviceStateAsync()
            {
                try
                {
                    var info = await deviceClient.GetDeviceInfoAsync(deviceKey).ConfigureAwait(false);
                    state = info.DeviceOn;
                }
                catch
                {
                    // If we can't sync, assume off for safety
                    state = false;
                }
            }

            /// <summary>
            /// Asynchronously turns the Tapo device on.
            /// </summary>
            public async Task TurnOnAsync()
            {
                await ExecuteWithReconnectAsync(async () =>
                {
                    await deviceClient.SetPowerAsync(deviceKey, true).ConfigureAwait(false);
                    state = true;
                }).ConfigureAwait(false);
            }

            /// <summary>
            /// Asynchronously turns the Tapo device off.
            /// </summary>
            public async Task TurnOffAsync()
            {
                await ExecuteWithReconnectAsync(async () =>
                {
                    await deviceClient.SetPowerAsync(deviceKey, false).ConfigureAwait(false);
                    state = false;
                }).ConfigureAwait(false);
            }

            /// <summary>
            /// Executes an action with automatic reconnection on network errors.
            /// Retries multiple times with exponential backoff.
            /// </summary>
            private async Task ExecuteWithReconnectAsync(Func<Task> action)
            {
                for (int attempt = 0; attempt < ConnectionRetryAttempts; attempt++)
                {
                    try
                    {
                        await action().ConfigureAwait(false);
                        isConnected = true;
                        InvokeOnUIThread(() =>
                        {
                            if (!isShuttingDown)
                            {
                                UpdateTooltip();
                            }
                        });
                        return;
                    }
                    catch (Exception ex) when (IsNetworkException(ex) && !isShuttingDown)
                    {
                        isConnected = false;
                        
                        // Last attempt failed
                        if (attempt == ConnectionRetryAttempts - 1)
                        {
                            InvokeOnUIThread(() =>
                            {
                                if (!isShuttingDown)
                                {
                                    UpdateTooltip("Connection lost. Will retry automatically.");
                                }
                            });
                            // Trigger background reconnection
                            _ = Task.Run(() => InitializeAsync());
                            throw;
                        }

                        // Wait before retry with exponential backoff capped at MaxRetryDelayMilliseconds
                        int delayMs = Math.Min(RetryDelayMilliseconds * (attempt + 1), MaxRetryDelayMilliseconds);
                        await Task.Delay(delayMs).ConfigureAwait(false);

                        try
                        {
                            await InitializeConnection().ConfigureAwait(false);
                        }
                        catch
                        {
                            // Will retry in next iteration
                        }
                    }
                    catch (Exception) when (!isShuttingDown)
                    {
                        // Non-network exception
                        InvokeOnUIThread(() =>
                        {
                            if (!isShuttingDown)
                            {
                                UpdateTooltip("Device error occurred.");
                            }
                        });
                        throw;
                    }
                }
            }

            /// <summary>
            /// Marshals an action to the UI thread.
            /// </summary>
            private void InvokeOnUIThread(Action action)
            {
                if (syncContext != null && !isShuttingDown)
                {
                    syncContext.Post(_ => action(), null);
                }
                else if (!isShuttingDown)
                {
                    action();
                }
            }

            /// <summary>
            /// Determines if an exception is network-related.
            /// </summary>
            private bool IsNetworkException(Exception ex)
            {
                return ex is HttpRequestException ||
                       ex is SocketException ||
                       ex is TaskCanceledException ||
                       ex is OperationCanceledException ||
                       (ex.InnerException != null && IsNetworkException(ex.InnerException));
            }

            /// <summary>
            /// Truncates tooltip text to Windows maximum length (63 characters).
            /// </summary>
            private string TruncateTooltip(string text)
            {
                if (string.IsNullOrEmpty(text))
                    return text;

                if (text.Length <= MaxTooltipLength)
                    return text;

                return text.Substring(0, MaxTooltipLength - 3) + "...";
            }

            /// <summary>
            /// Updates the tray icon based on current state.
            /// </summary>
            private void UpdateTrayIcon()
            {
                if (isShuttingDown || disposed)
                    return;

                trayIcon.Icon = state ? Resources.SwitchOn : Resources.SwitchOff;
            }

            /// <summary>
            /// Updates the tooltip with device information or status message.
            /// </summary>
            private void UpdateTooltip(string customMessage = null)
            {
                if (isShuttingDown || disposed)
                    return;

                string tooltipText;
                if (customMessage != null)
                {
                    tooltipText = customMessage;
                }
                else if (isConnected && deviceInfo != null)
                {
                    tooltipText = $"Click to toggle {deviceInfo.Model} ({deviceInfo.Nickname}) switch.";
                }
                else
                {
                    tooltipText = "Connecting to device...";
                }

                trayIcon.Text = TruncateTooltip(tooltipText);
            }

            /// <summary>
            /// Handles mouse click events on the tray icon to toggle the device state.
            /// </summary>
            /// <param name="sender">The event sender.</param>
            /// <param name="e">The mouse event arguments.</param>
            private async void MouseClick(object sender, MouseEventArgs e)
            {
                // Left mouse button toggles
                if (e.Button == MouseButtons.Left)
                {
                    if (!isConnected)
                    {
                        // Trigger immediate reconnection attempt
                        UpdateTooltip("Reconnecting...");
                        _ = Task.Run(() => InitializeAsync());
                        return;
                    }

                    try
                    {
                        if (state)
                        {
                            await TurnOffAsync().ConfigureAwait(false);
                        }
                        else
                        {
                            await TurnOnAsync().ConfigureAwait(false);
                        }

                        InvokeOnUIThread(() =>
                        {
                            if (!isShuttingDown)
                            {
                                UpdateTrayIcon();
                            }
                        });
                    }
                    catch
                    {
                        // Error already handled in ExecuteWithReconnectAsync
                    }
                }
            }

            /// <summary>
            /// Handles session ending events (shutdown, logoff) to turn off the device.
            /// This method temporarily cancels the shutdown to allow time for the device to turn off.
            /// </summary>
            private void OnSessionEnding(object sender, SessionEndingEventArgs e)
            {
                if (isShuttingDown)
                    return;

                isShuttingDown = true;

                // Cancel the shutdown temporarily to give us time to turn off the device
                e.Cancel = true;

                try
                {
                    // Use Task.Run().Wait() to ensure the async operation completes
                    // before allowing shutdown to proceed. Use timeout to prevent hanging.
                    var turnOffTask = Task.Run(async () =>
                    {
                        try
                        {
                            await TurnOffAsync().ConfigureAwait(false);
                        }
                        catch
                        {
                            // Silently fail
                        }
                    });

                    // Wait up to configured seconds for the device to turn off
                    if (!turnOffTask.Wait(TimeSpan.FromSeconds(ShutdownTimeoutSeconds)))
                    {
                        // Timeout - proceed with shutdown anyway
                    }
                }
                catch
                {
                    // Silently fail
                }
                finally
                {
                    // Now allow the shutdown to proceed
                    e.Cancel = false;
                    
                    // Clean up
                    Dispose();
                    
                    // If this is a logoff, we can exit cleanly
                    if (e.Reason == SessionEndReasons.Logoff)
                    {
                        Application.Exit();
                    }
                }
            }

            /// <summary>
            /// Handles session switch events (lock, unlock, remote connect/disconnect).
            /// </summary>
            private async void OnSessionSwitch(object sender, SessionSwitchEventArgs e)
            {
                switch (e.Reason)
                {
                    case SessionSwitchReason.SessionLogoff:
                        // Session logoff - turn off device
                        if (!isShuttingDown)
                        {
                            isShuttingDown = true;
                            try
                            {
                                await TurnOffAsync().ConfigureAwait(false);
                                InvokeOnUIThread(() =>
                                {
                                    if (!isShuttingDown)
                                    {
                                        UpdateTrayIcon();
                                    }
                                });
                            }
                            catch
                            {
                                // Silently fail
                            }
                        }
                        break;
                }
            }

            /// <summary>
            /// Handles application exit event to set shutdown flag.
            /// </summary>
            private void OnApplicationExit(object sender, EventArgs e)
            {
                isShuttingDown = true;
                Dispose();
            }

            /// <summary>
            /// Handles the exit menu item click event, turns off the device, and exits the application.
            /// </summary>
            /// <param name="sender">The event sender.</param>
            /// <param name="e">The event arguments.</param>
            private async void Exit(object sender, EventArgs e)
            {
                isShuttingDown = true;

                try
                {
                    await TurnOffAsync().ConfigureAwait(false);
                }
                catch
                {
                    // Silently fail on exit
                }

                if (!disposed)
                {
                    trayIcon.Visible = false;
                }
                Application.Exit();
            }

            /// <summary>
            /// Disposes of managed resources.
            /// </summary>
            public new void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            /// <summary>
            /// Disposes of managed resources.
            /// </summary>
            /// <param name="disposing">True if called from Dispose(), false if called from finalizer.</param>
            protected override void Dispose(bool disposing)
            {
                if (!disposed)
                {
                    if (disposing)
                    {
                        // Dispose managed resources
                        connectionCheckTimer?.Dispose();
                        connectionLock?.Dispose();
                        
                        if (trayIcon != null)
                        {
                            trayIcon.Visible = false;
                            trayIcon.Dispose();
                        }
                        
                        // Unregister session events
                        SystemEvents.SessionEnding -= OnSessionEnding;
                        SystemEvents.SessionSwitch -= OnSessionSwitch;
                        Application.ApplicationExit -= OnApplicationExit;
                    }

                    disposed = true;
                }

                base.Dispose(disposing);
            }
        }
    }
}
