using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Interop;

namespace GridCoords.Common
{
    internal static class Utils
    {
        // ── Win32 imports for multi-monitor window positioning ──

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MONITORINFO
        {
            public uint cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        private const uint MONITOR_DEFAULTTONEAREST = 2;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;

        internal static void PositionWindowCenterRight(Window window, IntPtr revitWindowHandle)
        {
            IntPtr hMonitor = MonitorFromWindow(revitWindowHandle, MONITOR_DEFAULTTONEAREST);

            var monitorInfo = new MONITORINFO();
            monitorInfo.cbSize = (uint)Marshal.SizeOf(typeof(MONITORINFO));

            if (!GetMonitorInfo(hMonitor, ref monitorInfo))
            {
                window.Left = SystemParameters.PrimaryScreenWidth - window.Width - 50;
                window.Top = (SystemParameters.PrimaryScreenHeight - window.Height) / 2;
                return;
            }

            RECT workArea = monitorInfo.rcWork;
            var wpfHelper = new WindowInteropHelper(window);
            IntPtr wpfHandle = wpfHelper.Handle;
            if (wpfHandle == IntPtr.Zero) return;

            GetWindowRect(wpfHandle, out RECT wpfRect);
            int wpfWidthPx = wpfRect.Right - wpfRect.Left;
            int wpfHeightPx = wpfRect.Bottom - wpfRect.Top;
            int workAreaHeight = workArea.Bottom - workArea.Top;

            int left = workArea.Right - wpfWidthPx - 50;
            int top = workArea.Top + (workAreaHeight - wpfHeightPx) / 2;

            SetWindowPos(wpfHandle, IntPtr.Zero, left, top, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
        }

        // ── Natural sort comparer for grid names ──

        internal static int NaturalSortCompare(string a, string b)
        {
            if (a == null && b == null) return 0;
            if (a == null) return -1;
            if (b == null) return 1;

            var partsA = Regex.Split(a, @"(\d+)");
            var partsB = Regex.Split(b, @"(\d+)");

            int len = Math.Min(partsA.Length, partsB.Length);
            for (int i = 0; i < len; i++)
            {
                int result;
                if (int.TryParse(partsA[i], out int numA) && int.TryParse(partsB[i], out int numB))
                    result = numA.CompareTo(numB);
                else
                    result = string.Compare(partsA[i], partsB[i], StringComparison.OrdinalIgnoreCase);

                if (result != 0) return result;
            }

            return partsA.Length.CompareTo(partsB.Length);
        }
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

            return curPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }
    }
}
