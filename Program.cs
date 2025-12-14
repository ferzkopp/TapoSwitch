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
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Set the default font for the application (example: Microsoft Sans Serif, 8.25pt)
            Application.SetDefaultFont(new Font("Microsoft Sans Serif", 8.25F));

            var context = new CustomApplicationContext();
            Application.Run(context);
        }

        /// <summary>
        /// Provides the application context for the TapoSwitch tray application, including tray icon and device control.
        /// </summary>
        public class CustomApplicationContext : ApplicationContext
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
            bool state = false;

            /// <summary>
            /// Indicates whether the application is shutting down.
            /// </summary>
            private bool isShuttingDown = false;

            /// <summary>
            /// Initializes a new instance of the <see cref="CustomApplicationContext"/> class.
            /// Sets up the tray icon and starts device initialization.
            /// </summary>
            public CustomApplicationContext()
            {
                // Initialize Tray Icon
                ContextMenuStrip contextMenuStrip = new ContextMenuStrip();
                contextMenuStrip.Items.Add("Exit", null, Exit);

                string tooltipText = 
                    deviceInfo == null ? "Click to toggle switch." : "Click to toggle " + deviceInfo.Model + " (" + deviceInfo.Nickname + ") switch.";
                trayIcon = new NotifyIcon()
                {
                    Icon = state ? Resources.SwitchOn : Resources.SwitchOff,
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
            }

            /// <summary>
            /// Asynchronously initializes the device connection and ensures the device is off at startup.
            /// </summary>
            private async Task InitializeAsync()
            {
                try
                {
                    await InitializeConnection();
                    await TurnOffAsync();
                }
                catch
                {
                    // Silently fail on initialization - user can try to toggle later
                }
            }

            /// <summary>
            /// Asynchronously establishes a connection to the Tapo device and retrieves device information.
            /// </summary>
            public async Task InitializeConnection()
            {
                deviceClient = new TapoDeviceClient(new List<ITapoDeviceClient>
                {
                    new SecurePassthroughDeviceClient(),
                    new KlapDeviceClient(),
                });

                deviceKey = await deviceClient.LoginByIpAsync(IpAddress, Username, Password);
                deviceInfo = await deviceClient.GetDeviceInfoAsync(deviceKey);
            }

            /// <summary>
            /// Asynchronously retrieves the latest device information.
            /// </summary>
            public async Task GetDeviceInfoAsync()
            {
                var deviceInfo = await deviceClient.GetDeviceInfoAsync(deviceKey);
            }

            /// <summary>
            /// Asynchronously turns the Tapo device on.
            /// </summary>
            public async Task TurnOnAsync()
            {
                await ExecuteWithReconnectAsync(() => deviceClient.SetPowerAsync(deviceKey, true));
            }

            /// <summary>
            /// Asynchronously turns the Tapo device off.
            /// </summary>
            public async Task TurnOffAsync()
            {
                await ExecuteWithReconnectAsync(() => deviceClient.SetPowerAsync(deviceKey, false));
            }

            /// <summary>
            /// Executes an action with automatic reconnection on network errors.
            /// </summary>
            private async Task ExecuteWithReconnectAsync(Func<Task> action)
            {
                try
                {
                    await action();
                }
                catch (Exception ex) when (IsNetworkException(ex) && !isShuttingDown)
                {
                    try
                    {
                        await InitializeConnection();
                        await action();
                    }
                    catch
                    {
                        // Silently fail after reconnection attempt
                    }
                }
                catch
                {
                    // Silently fail for non-network exceptions
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
                       ex is OperationCanceledException;
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
                    state = !state;
                    if (state)
                    {
                        await TurnOnAsync();
                    }
                    else
                    {
                        await TurnOffAsync();
                    }

                    trayIcon.Icon = state ? Resources.SwitchOn : Resources.SwitchOff;
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
                            await TurnOffAsync();
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
                    trayIcon.Visible = false;
                    
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
                                await TurnOffAsync();
                                state = false;
                                trayIcon.Icon = Resources.SwitchOff;
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
                
                // Unregister session events
                SystemEvents.SessionEnding -= OnSessionEnding;
                SystemEvents.SessionSwitch -= OnSessionSwitch;
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
                    await TurnOffAsync();
                }
                catch
                {
                    // Silently fail on exit
                }

                trayIcon.Visible = false;
                Application.Exit();
            }
        }
    }
}

































