using ReportManager.Addon.Logging;
using ReportManager.Addon.Screens;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Core
{
    public sealed class AddonBootstrap
    {
        private readonly SapApplication _sap;
        private readonly Logger _log;

        public AddonBootstrap(SapApplication sap, Logger log)
        {
            _sap = sap ?? throw new ArgumentNullException(nameof(sap));
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        public void Start()
        {
            _log.Info("Iniciando Add-On...");

            // 1) Cargar SRF
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string srfPath = Path.Combine(baseDir, "Forms", "Principal.srf");

            var loader = new SrfFormLoader(_sap.App);
            loader.LoadFromFile(srfPath);
            _log.Info("SRF cargado: " + srfPath);

            // 2) Enlazar eventos del screen
            var principal = new PrincipalScreen(_sap.App, _log);
            principal.WireEvents();

            _sap.App.StatusBar.SetText("Add-On ReportManager cargado.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            _log.Info("Add-On listo.");
        }
    }
}
