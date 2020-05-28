using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadSelectedText
{
    /// <summary>
    /// Reads the selected text from most windows.
    /// </summary>
    /// <example>
    /// WindowInteropHelper nativeWindow = new WindowInteropHelper(this);
    /// HwndSource hwndSource = HwndSource.FromHwnd(nativeWindow.Handle);
    /// SelectionReader selectionReader = new SelectionReader(nativeWindow.Handle);
    /// hwndSource.AddHook(selectionReader.WindowProc);
    /// </example>
    public class SelectionReader
    {
        private readonly int shellMessage;

        private IntPtr activeWindow;
        private IntPtr lastWindow;

        /// <summary>Set when the active window has changed.</summary>
        private readonly AutoResetEvent gotActiveWindow = new AutoResetEvent(false);
        /// <summary>Set when the clipboard has been updated.</summary>
        private readonly AutoResetEvent gotClipboard = new AutoResetEvent(false);

        public SelectionReader(IntPtr hwnd)
        {
            this.shellMessage = WinApi.RegisterWindowMessage("SHELLHOOK");
            WinApi.RegisterShellHookWindow(hwnd);
            WinApi.AddClipboardFormatListener(hwnd);
        }

        /// <summary>
        /// Gets the selected text of the given window, or the last activate window.
        /// </summary>
        /// <param name="windowHandle">The window.</param>
        /// <returns>The selected text.</returns>
        public async Task<string> GetSelectedText(IntPtr? windowHandle = null)
        {
            IntPtr hwnd = windowHandle ?? this.lastWindow;
            await Task.Run(() =>
            {
                if (hwnd != IntPtr.Zero)
                {
                    // Activate the window, if it's not already.
                    IntPtr active = WinApi.GetForegroundWindow();
                    if (active != hwnd)
                    {
                        this.gotActiveWindow.Reset();
                        WinApi.SetForegroundWindow(hwnd);

                        // Wait it to be activated.
                        this.gotActiveWindow.WaitOne(3000);
                    }
                }

                // Copy the selected text to clipboard
                this.gotClipboard.Reset();
                SendKeys.SendWait("^c");

                // Wait for the clipboard update.
                this.gotClipboard.WaitOne(3000);
            });

            return System.Windows.Clipboard.GetText();
        }

        /// <summary>
        /// A Window procedure - handles window messages for a window.
        /// Pass this method to HwndSource.AddHook().
        /// </summary>
        /// <param name="hwnd">The window handle.</param>
        /// <param name="msg">The message.</param>
        /// <param name="wParam">Message data.</param>
        /// <param name="lParam">Message data.</param>
        /// <param name="handled">Set to true if the message has been handled</param>
        /// <returns>The message result.</returns>
        public IntPtr WindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            switch (msg)
            {
                case WinApi.WM_CLIPBOARDUPDATE:
                    // The clipboard has been updated.
                    this.gotClipboard.Set();
                    handled = true;
                    break;

                default:
                    if (msg == this.shellMessage)
                    {
                        // A window has been activated.
                        if (wParam.ToInt32() == WinApi.HSHELL_WINDOWACTIVATED
                            || wParam.ToInt32() == WinApi.HSHELL_RUDEAPPACTIVATED)
                        {
                            // The activated window is passed via lParam, but this wasn't accurate
                            // for Modern UI apps.
                            IntPtr window = WinApi.GetForegroundWindow();
                            if (this.activeWindow != window)
                            {
                                this.activeWindow = window;
                                // Ignore the application window
                                // TODO: Get the window's process, and check against this process.
                                if (window != hwnd)
                                {
                                    this.lastWindow = window;
                                }

                                this.gotActiveWindow.Set();
                            }
                        }
                    }
                    break;
            }

            return IntPtr.Zero;
        }
    }

    internal static class WinApi
    {
        public const int HSHELL_WINDOWACTIVATED = 4;
        public const int HSHELL_RUDEAPPACTIVATED = 0x8004;

        public const int WM_CLIPBOARDUPDATE = 0x031D;
        public static IntPtr HWND_MESSAGE = new IntPtr(-3);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterShellHookWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int RegisterWindowMessage(string lpString);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
        
        public static short HighWord(this IntPtr n) => n.ToInt32().HighWord();
        public static short LowWord(this IntPtr n) => n.ToInt32().LowWord();

        public static short HighWord(this int n)
        {
            return ((short)(n >> 16));
        }

        public static short LowWord(this int n)
        {
            return ((short)(n & 0xffff));
        }
    }
}
