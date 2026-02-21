using ReportManager.Addon.Core;
using ReportManager.Addon.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportManager.Addon
{
    internal static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "logAddonRM"
            );

            var log = new Logger(logDir);

            try
            {
                // SAP B1 pasa el connection string en args[0]
                string connStr = (args != null && args.Length > 0) ? args[0] : null;

                var sap = SapApplication.Connect(connStr);
                var bootstrap = new AddonBootstrap(sap, log);
                bootstrap.Start();

                // Mantener vivo el proceso
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                log.Error("Fallo crítico al iniciar el Add-On", ex);

                // Si no hay SAP App, al menos muestra algo
                try { System.Windows.Forms.MessageBox.Show(ex.ToString(), "Add-On Error"); } catch { }
            }
        }
    }
}
