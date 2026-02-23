using ReportManager.Addon.Logging;
using SAPbouiCOM;
using System;
using System.Threading;
using System.Windows.Forms;

namespace ReportManager.Addon.Core.Embedding
{
    public sealed class EmbeddedWinFormHost : IDisposable
    {
        private readonly SAPbouiCOM.Application _app;
        private readonly Logger _log;

        public string HostSapFormUid { get; }
        public string HostTitle { get; }

        private SAPbouiCOM.Form _sapHostForm;
        private System.Windows.Forms.Form _childWinForm;
        private Thread _uiThread;

        public EmbeddedWinFormHost(SAPbouiCOM.Application app, Logger log, string hostSapFormUid, string hostTitle)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _log = log ?? throw new ArgumentNullException(nameof(log));
            HostSapFormUid = hostSapFormUid ?? throw new ArgumentNullException(nameof(hostSapFormUid));
            HostTitle = hostTitle ?? throw new ArgumentNullException(nameof(hostTitle));
        }

        public void ShowOrFocus<TWinForm>(Func<TWinForm> factory, int width, int height)
            where TWinForm : System.Windows.Forms.Form
        {
            // Si ya está vivo, solo enfoca el SAP host
            if (IsAlive())
            {
                TryFocusHost();
                return;
            }

            var sapMainHwnd = SapWindowHandles.GetSapMainHwnd(_app);
            if (sapMainHwnd == IntPtr.Zero)
                throw new InvalidOperationException("No se pudo obtener el HWND principal de SAP Business One.");

            _sapHostForm = SapHostFormFactory.CreateOrFocusHostForm(_app, HostSapFormUid, HostTitle, width, height);
            var hostHwnd = SapHostFormFactory.GetHostFormHwnd(_app, _sapHostForm, sapMainHwnd);

            if (hostHwnd == IntPtr.Zero)
                throw new InvalidOperationException("No se pudo obtener el HWND del formulario host SAP para incrustar.");

            _uiThread = new Thread(() =>
            {
                try
                {
                    var wf = factory();
                    _childWinForm = wf;

                    // Importante para embedding estable
                    wf.TopLevel = false;
                    wf.FormBorderStyle = FormBorderStyle.None;
                    wf.Dock = DockStyle.Fill;

                    // Parent + estilo WS_CHILD y sin bordes
                    Win32Native.SetParent(wf.Handle, hostHwnd);

                    int style = Win32Native.GetWindowLong(wf.Handle, Win32Native.GWL_STYLE);
                    style |= Win32Native.WS_CHILD;

                    style &= ~Win32Native.WS_CAPTION;
                    style &= ~Win32Native.WS_THICKFRAME;
                    style &= ~Win32Native.WS_MINIMIZEBOX;
                    style &= ~Win32Native.WS_MAXIMIZEBOX;
                    style &= ~Win32Native.WS_SYSMENU;

                    Win32Native.SetWindowLong(wf.Handle, Win32Native.GWL_STYLE, style);

                    wf.Show();
                    System.Windows.Forms.Application.Run(wf);
                }
                catch (Exception ex)
                {
                    try { _log.Error("Error al incrustar WinForm.", ex); } catch { }
                }
            });

            _uiThread.IsBackground = true;
            _uiThread.SetApartmentState(ApartmentState.STA);
            _uiThread.Start();

            TryFocusHost();
        }

        private bool IsAlive()
        {
            try
            {
                if (_sapHostForm == null) return false;
                // Si el form SAP ya no existe, se considera muerto.
                var check = _app.Forms.Item(HostSapFormUid);
                if (check == null) return false;

                return _childWinForm != null && !_childWinForm.IsDisposed;
            }
            catch
            {
                return false;
            }
        }

        private void TryFocusHost()
        {
            try
            {
                var f = _app.Forms.Item(HostSapFormUid);
                f.Visible = true;
                f.Select();
            }
            catch { }
        }

        public void Dispose()
        {
            try
            {
                if (_childWinForm != null && !_childWinForm.IsDisposed)
                {
                    try
                    {
                        // Cierra el loop del thread STA
                        _childWinForm.Invoke(new Action(() => _childWinForm.Close()));
                    }
                    catch
                    {
                        try { _childWinForm.Close(); } catch { }
                    }
                }
            }
            catch { }

            _childWinForm = null;

            // No intentamos abortar thread; cerrando el form se libera el Application.Run.
            _uiThread = null;
            _sapHostForm = null;
        }
    }
}
