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
        private readonly Application _app;
        private readonly Logger _log;

        public PrincipalScreen(Application app, Logger log)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        public void WireEvents()
        {
            _app.ItemEvent += OnItemEvent;
        }

        private void OnItemEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                // Filtrar SOLO nuestro form
                if (formUID != FormUid)
                    return;

                // Botón presionado (AfterAction)
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
                // No rompemos SAP; solo notificamos
                _app.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
