using ReportManager.Addon.Logging;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WinForms = System.Windows.Forms;

namespace ReportManager.Addon.Services
{
    /// <summary>
    /// Servicio reutilizable para incrustar formularios WinForms dentro de un formulario de SAP Business One.
    /// </summary>
    public sealed class SapWinFormEmbedder
    {
        private const int GwlStyle = -16;
        private const int WsChild = 0x40000000;
        private const int WsVisible = 0x10000000;
        private const int WsPopup = unchecked((int)0x80000000);
        private const uint SwpNoZOrder = 0x0004;
        private const uint SwpNoActivate = 0x0010;

        private readonly Logger _log;
        private readonly Dictionary<string, WinForms.Form> _embeddedForms = new Dictionary<string, WinForms.Form>(StringComparer.OrdinalIgnoreCase);

        public SapWinFormEmbedder(Logger log)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        public void ShowOrEmbed<TForm>(Form sapForm, string instanceKey, Func<TForm> formFactory, int left, int top, int width, int height)
            where TForm : WinForms.Form
        {
            if (sapForm == null)
            {
                throw new ArgumentNullException(nameof(sapForm));
            }

            if (string.IsNullOrWhiteSpace(instanceKey))
            {
                throw new ArgumentException("Se requiere una clave de instancia.", nameof(instanceKey));
            }

            if (formFactory == null)
            {
                throw new ArgumentNullException(nameof(formFactory));
            }

            CleanupDisposedForms();

            if (!_embeddedForms.TryGetValue(instanceKey, out var winForm) || winForm.IsDisposed)
            {
                winForm = formFactory();
                _embeddedForms[instanceKey] = winForm;
            }

            var parentHandle = ResolveSapFormHandle(sapForm.Title);
            if (parentHandle == IntPtr.Zero)
            {
                _log.Warn($"No se pudo resolver el HWND del formulario SAP '{sapForm.Title}'. Se mostrará el formulario WinForms en modo flotante.");
                ShowFloating(winForm);
                return;
            }

            if (!winForm.Visible)
            {
                winForm.Show();
            }

            winForm.TopLevel = false;
            winForm.FormBorderStyle = WinForms.FormBorderStyle.None;

            var childHandle = winForm.Handle;
            SetParent(childHandle, parentHandle);

            var currentStyle = GetWindowLong(childHandle, GwlStyle);
            var newStyle = (currentStyle | WsChild | WsVisible) & ~WsPopup;
            SetWindowLong(childHandle, GwlStyle, newStyle);

            SetWindowPos(childHandle, IntPtr.Zero, left, top, width, height, SwpNoActivate | SwpNoZOrder);
            winForm.BringToFront();
        }

        private static void ShowFloating(WinForms.Form winForm)
        {
            winForm.TopLevel = true;
            winForm.FormBorderStyle = WinForms.FormBorderStyle.Sizable;
            winForm.StartPosition = WinForms.FormStartPosition.CenterScreen;
            winForm.Show();
            winForm.BringToFront();
        }

        private void CleanupDisposedForms()
        {
            var disposedKeys = new List<string>();
            foreach (var pair in _embeddedForms)
            {
                if (pair.Value == null || pair.Value.IsDisposed)
                {
                    disposedKeys.Add(pair.Key);
                }
            }

            for (int i = 0; i < disposedKeys.Count; i++)
            {
                _embeddedForms.Remove(disposedKeys[i]);
            }
        }

        private static IntPtr ResolveSapFormHandle(string sapFormTitle)
        {
            if (string.IsNullOrWhiteSpace(sapFormTitle))
            {
                return IntPtr.Zero;
            }

            for (int i = 0; i < 8; i++)
            {
                var handle = FindWindow(null, sapFormTitle);
                if (handle != IntPtr.Zero)
                {
                    return handle;
                }

                Thread.Sleep(40);
            }

            return IntPtr.Zero;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);
    }

}
