using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowCenterTray
{
    public partial class TrayApplication : Form
    {
        private NotifyIcon? trayIcon;
        private ContextMenuStrip? contextMenu;
        private GlobalHotkey? centerHotkey;
        private GlobalHotkey? alwaysOnTopHotkey;

        // Windows API Imports
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOMOVE = 0x0002;
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOPMOST = 0x00000008;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public TrayApplication()
        {
            InitializeComponent();
            CreateTrayIcon();
            RegisterHotkeys();
            
            // Fenster verstecken
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Visible = false;
        }

        private void InitializeComponent()
        {
            this.Text = "Window Center Tool";
            this.Size = new Size(1, 1);
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
        }

        private void CreateTrayIcon()
        {
            // Erstelle ein einfaches Icon (16x16 weißer Kreis auf transparentem Hintergrund)
            Bitmap iconBitmap = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(iconBitmap))
            {
                g.Clear(Color.Transparent);
                g.FillEllipse(Brushes.White, 2, 2, 12, 12);
                g.DrawEllipse(Pens.Black, 2, 2, 12, 12);
            }

            // Erstelle Kontext-Menü
            contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Center Window (Ctrl+Alt+C)", null, CenterCurrentWindow);
            contextMenu.Items.Add("Toggle Always On Top (Alt+Space)", null, ToggleAlwaysOnTop);
            contextMenu.Items.Add("-"); // Separator
            contextMenu.Items.Add("Exit", null, ExitApplication);

            // Erstelle Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Icon.FromHandle(iconBitmap.GetHicon()),
                ContextMenuStrip = contextMenu,
                Visible = true,
                Text = "Window Center Tool - Ctrl+Alt+C to center, Alt+Space for always on top"
            };

            trayIcon.DoubleClick += TrayIcon_DoubleClick;
        }

        private void RegisterHotkeys()
        {
            // Registriere Ctrl+Alt+C als globalen Hotkey für Zentrierung
            centerHotkey = new GlobalHotkey(Keys.C, KeyModifiers.Control | KeyModifiers.Alt, this);
            if (!centerHotkey.Register())
            {
                MessageBox.Show("Could not register hotkey Ctrl+Alt+C. The key combination might already be in use.",
                    "Hotkey Registration Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // Registriere Alt+Space als globalen Hotkey für Always On Top
            alwaysOnTopHotkey = new GlobalHotkey(Keys.Space, KeyModifiers.Alt, this);
            if (!alwaysOnTopHotkey.Register())
            {
                MessageBox.Show("Could not register hotkey Alt+Space. The key combination might already be in use.",
                    "Hotkey Registration Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void TrayIcon_DoubleClick(object? sender, EventArgs e)
        {
            CenterCurrentWindow(sender!, e);
        }

        private void CenterCurrentWindow(object? sender, EventArgs e)
        {
            IntPtr activeWindow = GetForegroundWindow();
            
            if (activeWindow == IntPtr.Zero || !IsWindow(activeWindow) || !IsWindowVisible(activeWindow))
            {
                ShowBalloonTip("No active window found", ToolTipIcon.Warning);
                return;
            }

            if (GetWindowRect(activeWindow, out RECT windowRect))
            {
                // Berechne Fenstergröße
                int windowWidth = windowRect.Right - windowRect.Left;
                int windowHeight = windowRect.Bottom - windowRect.Top;

                // Bekomme Bildschirmauflösung
                Screen currentScreen = Screen.FromHandle(activeWindow);
                Rectangle screenBounds = currentScreen.WorkingArea;

                // Berechne zentrierte Position
                int centerX = screenBounds.X + (screenBounds.Width - windowWidth) / 2;
                int centerY = screenBounds.Y + (screenBounds.Height - windowHeight) / 2;

                // Setze Fenster an zentrierte Position
                if (SetWindowPos(activeWindow, IntPtr.Zero, centerX, centerY, 0, 0, 
                    SWP_NOSIZE | SWP_NOZORDER))
                {
                    ShowBalloonTip("Window centered!", ToolTipIcon.Info);
                }
                else
                {
                    ShowBalloonTip("Failed to center window", ToolTipIcon.Error);
                }
            }
        }

        private void ToggleAlwaysOnTop(object? sender, EventArgs e)
        {
            IntPtr activeWindow = GetForegroundWindow();
            
            if (activeWindow == IntPtr.Zero || !IsWindow(activeWindow) || !IsWindowVisible(activeWindow))
            {
                ShowBalloonTip("No active window found", ToolTipIcon.Warning);
                return;
            }

            // Prüfe aktuellen Always-On-Top Status
            IntPtr exStyle = GetWindowLong(activeWindow, GWL_EXSTYLE);
            bool isTopmost = (exStyle.ToInt32() & WS_EX_TOPMOST) == WS_EX_TOPMOST;

            if (isTopmost)
            {
                // Always-On-Top ausschalten
                if (SetWindowPos(activeWindow, HWND_NOTOPMOST, 0, 0, 0, 0, 
                    SWP_NOSIZE | SWP_NOMOVE))
                {
                    ShowBalloonTip("Always On Top disabled", ToolTipIcon.Info);
                }
                else
                {
                    ShowBalloonTip("Failed to disable Always On Top", ToolTipIcon.Error);
                }
            }
            else
            {
                // Always-On-Top einschalten
                if (SetWindowPos(activeWindow, HWND_TOPMOST, 0, 0, 0, 0, 
                    SWP_NOSIZE | SWP_NOMOVE))
                {
                    ShowBalloonTip("Always On Top enabled", ToolTipIcon.Info);
                }
                else
                {
                    ShowBalloonTip("Failed to enable Always On Top", ToolTipIcon.Error);
                }
            }
        }

        private void ShowBalloonTip(string message, ToolTipIcon icon)
        {
            trayIcon!.ShowBalloonTip(2000, "Window Center Tool", message, icon);
        }

        private void ExitApplication(object? sender, EventArgs e)
        {
            centerHotkey?.Unregister();
            alwaysOnTopHotkey?.Unregister();
            trayIcon!.Visible = false;
            Application.Exit();
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(false);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
            else
            {
                centerHotkey?.Unregister();
                alwaysOnTopHotkey?.Unregister();
                trayIcon?.Dispose();
                base.OnFormClosing(e);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == GlobalHotkey.WM_HOTKEY)
            {
                // Bestimme welcher Hotkey gedrückt wurde
                int hotkeyId = m.WParam.ToInt32();
                
                if (centerHotkey != null && hotkeyId == centerHotkey.Id)
                {
                    CenterCurrentWindow(this, EventArgs.Empty);
                }
                else if (alwaysOnTopHotkey != null && hotkeyId == alwaysOnTopHotkey.Id)
                {
                    ToggleAlwaysOnTop(this, EventArgs.Empty);
                }
            }
            base.WndProc(ref m);
        }
    }

    // Globaler Hotkey Handler
    public class GlobalHotkey
    {
        public const int WM_HOTKEY = 0x0312;
        private static int currentId = 0;
        private IntPtr hWnd;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public GlobalHotkey(Keys key, KeyModifiers modifiers, Form form)
        {
            this.Id = System.Threading.Interlocked.Increment(ref currentId);
            this.hWnd = form.Handle;
            this.Key = key;
            this.Modifiers = modifiers;
        }

        public int Id { get; private set; }
        public Keys Key { get; private set; }
        public KeyModifiers Modifiers { get; private set; }

        public bool Register()
        {
            return RegisterHotKey(hWnd, Id, (uint)Modifiers, (uint)Key);
        }

        public bool Unregister()
        {
            return UnregisterHotKey(hWnd, Id);
        }
    }

    [Flags]
    public enum KeyModifiers
    {
        None = 0,
        Alt = 1,
        Control = 2,
        Shift = 4,
        Windows = 8
    }

    // Program Entry Point
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            using (var trayApp = new TrayApplication())
            {
                Application.Run();
            }
        }
    }
}