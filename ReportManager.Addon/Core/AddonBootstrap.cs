using ReportManager.Addon.Entidades;
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
        private const string SapTopMenuId = "43520";
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

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string srfPath = Path.Combine(baseDir, "Forms", "Principal.srf");

            var loader = new SrfFormLoader(_sap.App);
            var menuManager = new MenuManager(_sap.App);
            menuManager.EnsurePopupWithEntry(
                SapTopMenuId,
                PrincipalScreen.PopupMenuId,
                "ReportManager",
                PrincipalScreen.OpenPrincipalMenuId,
                "Principal");
            menuManager.EnsurePopupWithEntry(
                SapTopMenuId,
                PrincipalScreen.PopupMenuId,
                "ReportManager",
                PrincipalScreen.OpenConfigMenuId,
                "Configuración");
            _log.Info("Menú ReportManager registrado (Principal y Configuración).");

            var principalFormController = new PrincipalFormController(_sap.App, loader, srfPath, PrincipalScreen.FormUid);
            var configMetadataService = new ReportManager.Addon.Services.ConfigurationMetadataService(_sap.App, _log, Globals.rCompany);
            var principal = new PrincipalScreen(_sap.App, _log, principalFormController, configMetadataService);
            principal.WireEvents();

            _sap.App.StatusBar.SetText("Add-On ReportManager cargado.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            _log.Info("Add-On listo.");
        }
    }
}
