using ReportManager.Addon.Core;
using ReportManager.Addon.Logging;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Screens
{
    public sealed class PrincipalScreen
    {
        public const string FormUid = "RM_PRINCIPAL";
        public const string PopupMenuId = "RM_MENU";
        public const string OpenPrincipalMenuId = "RM_MENU_PRINCIPAL";
        private readonly Application _app;
        private readonly Logger _log;
        private readonly PrincipalFormController _principalFormController;

        public PrincipalScreen(Application app, Logger log, PrincipalFormController principalFormController)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _principalFormController = principalFormController ?? throw new ArgumentNullException(nameof(principalFormController));
        }

        public void WireEvents()
        {
            _app.MenuEvent += OnMenuEvent;
            _app.ItemEvent += OnItemEvent;
        }

        private void OnMenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                if (!pVal.BeforeAction && pVal.MenuUID == OpenPrincipalMenuId)
                {
                    _principalFormController.OpenOrFocus();
                    _log.Info("Menú Principal ejecutado.");
                }
            }
            catch (Exception ex)
            {
                _log.Error("Error en OnMenuEvent", ex);
                _app.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void OnItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                if (formUID != FormUid)
                    return;

                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED
                    && pVal.ItemUID == "btnHello"
                    && pVal.ActionSuccess)
                {
                    _app.StatusBar.SetText("Se ha presionado el botón.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    _app.MessageBox("Se ha presionado el botón.");
                    _log.Info("btnHello presionado en RM_PRINCIPAL");
                }
            }
            catch (Exception ex)
            {
                _log.Error("Error en OnItemEvent", ex);
                _app.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
