using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace QuickLook.Plugin.OfficeViewer;

internal static class SyncfusionTrial
{
    private static class User32
    {
        public const int SW_HIDE = 0;
        public const uint SWP_NOZORDER = 0x0004;
        public const uint SWP_NOMOVE = 0x0002;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ShowWindow(nint hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool DestroyWindow(nint hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool EnableWindow(nint hWnd, bool enable);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(nint hWnd, nint hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    }

    private static bool _trialed = false;

    public static void Start(Window mainWindow)
    {
        if (_trialed) return;

        _ = Task.Run(() =>
        {
            for (int i = 0; i < 9999; i++)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    Window window = Application.Current.Windows.OfType<Window>()
                        .FirstOrDefault(w => w.Title.StartsWith("Syncfusion"));

                    if (window != null)
                    {
                        _ = User32.EnableWindow(new WindowInteropHelper(mainWindow).Handle, true);

                        try
                        {
                            if (window.IsVisible)
                            {
                                nint hWnd = new WindowInteropHelper(window).Handle;
                                window.WindowStyle = WindowStyle.None;
                                window.ResizeMode = ResizeMode.NoResize;
                                _ = User32.SetWindowPos(hWnd, IntPtr.Zero, -90000, -90000, 0, 0, User32.SWP_NOZORDER | User32.SWP_NOMOVE);
                                _ = User32.ShowWindow(hWnd, User32.SW_HIDE);
                                window.Close();
                                _ = User32.DestroyWindow(hWnd);
                                mainWindow?.Activate();
                                mainWindow?.Focus();
                                _trialed = true;
                            }
                        }
                        catch { }
                    }
                });

                if (_trialed == true) break;
            }
        });
    }
}
