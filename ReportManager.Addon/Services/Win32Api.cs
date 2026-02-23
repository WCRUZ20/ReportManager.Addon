using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Services
{
    internal static class Win32Api
    {
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        internal delegate bool EnumWindowsDelegate(IntPtr hWnd, int lParam);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int EnumChildWindows(IntPtr hWndParent, EnumWindowsDelegate lpEnumFunc, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int cch);
    }
}
