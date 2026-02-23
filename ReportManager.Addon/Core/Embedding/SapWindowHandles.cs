using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace ReportManager.Addon.Core.Embedding
{
    internal static class SapWindowHandles
    {
        public static IntPtr GetSapMainHwnd(Application app)
        {
            // Igual enfoque que la guía VB: compara MainWindowTitle con app.Desktop.Title
            foreach (var p in Process.GetProcessesByName("SAP Business One"))
            {
                try
                {
                    if (p.MainWindowHandle == IntPtr.Zero) continue;
                    if (string.Equals(p.MainWindowTitle, app.Desktop.Title, StringComparison.OrdinalIgnoreCase))
                        return p.MainWindowHandle;
                }
                catch { /* ignore */ }
            }

            // Fallback: primer handle válido
            foreach (var p in Process.GetProcessesByName("SAP Business One"))
            {
                try
                {
                    if (p.MainWindowHandle != IntPtr.Zero)
                        return p.MainWindowHandle;
                }
                catch { /* ignore */ }
            }

            return IntPtr.Zero;
        }

        public static IntPtr FindChildWindowByExactTitle(IntPtr parentHwnd, string exactTitle)
        {
            if (parentHwnd == IntPtr.Zero || string.IsNullOrWhiteSpace(exactTitle))
                return IntPtr.Zero;

            var matches = new List<IntPtr>();

            Win32Native.EnumChildWindows(parentHwnd, (hWnd, lParam) =>
            {
                var sb = new StringBuilder(512);
                int len = Win32Native.GetWindowText(hWnd, sb, sb.Capacity);
                if (len > 0)
                {
                    var title = sb.ToString(0, len);
                    if (string.Equals(title, exactTitle, StringComparison.OrdinalIgnoreCase))
                        matches.Add(hWnd);
                }
                return true;
            }, 0);

            return matches.Count > 0 ? matches[0] : IntPtr.Zero;
        }
    }
}
