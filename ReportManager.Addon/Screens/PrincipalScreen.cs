using ReportManager.Addon.Core;
using ReportManager.Addon.Entidades;
using ReportManager.Addon.Logging;
using ReportManager.Addon.Services;
using SAPbobsCOM;
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
        public const string OpenConfigMenuId = "RM_MENU_CONFIG";

        private const string LoginFormUid = "RM_CFG_LOGIN";
        private const string ConfigFormUid = "RM_CFG_FORM";
        private const string DepartmentEditUid = "edt_dpt";
        private const string ReportEditUid = "edt_rpt";
        private const string DepartmentChooseFromListUid = "RM_CFL_DPT";
        private const string ReportChooseFromListUid = "RM_CFL_RPT";
        private readonly Application _app;
        private readonly Logger _log;
        private readonly PrincipalFormController _principalFormController;
        private readonly ConfigurationMetadataService _configurationMetadataService;

        public PrincipalScreen(
            Application app,
            Logger log,
            PrincipalFormController principalFormController,
            ConfigurationMetadataService configurationMetadataService)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _principalFormController = principalFormController ?? throw new ArgumentNullException(nameof(principalFormController));
            _configurationMetadataService = configurationMetadataService ?? throw new ArgumentNullException(nameof(configurationMetadataService));
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

                if (!pVal.BeforeAction && pVal.MenuUID == OpenConfigMenuId)
                {
                    OpenConfigurationLoginForm();
                    _log.Info("Menú Configuración ejecutado.");
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
                if (formUID == FormUid
                    && pVal.EventType == BoEventTypes.et_FORM_LOAD
                    && pVal.ActionSuccess)
                {
                    ConfigureChooseFromLists();
                    return;
                }

                if (formUID == FormUid
                    && pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == ReportEditUid)
                {
                    ApplyReportChooseFromListFilter();
                    return;
                }

                if (formUID == FormUid
                    && pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction)
                {
                    HandleChooseFromListSelection(ref pVal);
                    return;
                }

                if (formUID == FormUid
                    && pVal.EventType == BoEventTypes.et_ITEM_PRESSED
                    && pVal.ItemUID == "btn_exe"
                    && pVal.ActionSuccess)
                {
                    _app.StatusBar.SetText("Se ha presionado el botón.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    _log.Info("btn_exe presionado en RM_PRINCIPAL");
                    return;
                }

                if (formUID == LoginFormUid
                    && pVal.EventType == BoEventTypes.et_ITEM_PRESSED
                    && pVal.ItemUID == "btn_ok"
                    && pVal.ActionSuccess)
                {
                    ValidateCredentialsAndOpenConfiguration();
                    return;
                }

                if (formUID == ConfigFormUid
                    && pVal.EventType == BoEventTypes.et_ITEM_PRESSED
                    && pVal.ActionSuccess)
                {
                    if (pVal.ItemUID == "btn_dpts")
                    {
                        _configurationMetadataService.CreateDepartmentTable();
                        _configurationMetadataService.CreateParamterTypeTable();
                        return;
                    }

                    if (pVal.ItemUID == "btn_udos")
                    {
                        _configurationMetadataService.CreateReportConfigurationStructures();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                _log.Error("Error en OnItemEvent", ex);
                _app.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void ConfigureChooseFromLists()
        {
            var form = TryGetOpenForm(FormUid);
            if (form == null)
            {
                return;
            }

            EnsureChooseFromList(form, DepartmentChooseFromListUid, "SS_DPTS");
            EnsureChooseFromList(form, ReportChooseFromListUid, "SS_PRM_CAB");

            var departmentEdit = (EditText)form.Items.Item(DepartmentEditUid).Specific;
            departmentEdit.ChooseFromListUID = DepartmentChooseFromListUid;
            departmentEdit.ChooseFromListAlias = "Code";

            var reportEdit = (EditText)form.Items.Item(ReportEditUid).Specific;
            reportEdit.ChooseFromListUID = ReportChooseFromListUid;
            reportEdit.ChooseFromListAlias = "U_SS_IDRPT";
        }

        private void EnsureChooseFromList(Form form, string chooseFromListUid, string objectType)
        {
            try
            {
                if (form.ChooseFromLists.Item(chooseFromListUid) != null)
                {
                    return;
                }
            }
            catch
            {
            }

            var chooseFromListParams = (ChooseFromListCreationParams)_app.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams);
            chooseFromListParams.UniqueID = chooseFromListUid;
            chooseFromListParams.ObjectType = objectType;
            chooseFromListParams.MultiSelection = false;
            form.ChooseFromLists.Add(chooseFromListParams);
        }

        private void ApplyReportChooseFromListFilter()
        {
            var form = TryGetOpenForm(FormUid);
            if (form == null)
            {
                return;
            }

            var departmentValue = ((EditText)form.Items.Item(DepartmentEditUid).Specific).Value?.Trim();
            var reportChooseFromList = form.ChooseFromLists.Item(ReportChooseFromListUid);
            var conditions = (Conditions)_app.CreateObject(BoCreatableObjectType.cot_Conditions);

            if (string.IsNullOrWhiteSpace(departmentValue))
            {
                reportChooseFromList.SetConditions(conditions);
                return;
            }

            var reportIds = GetReportIdsByDepartment(departmentValue);
            if (reportIds.Count == 0)
            {
                var noDataCondition = conditions.Add();
                noDataCondition.Alias = "Code";
                noDataCondition.Operation = BoConditionOperation.co_EQUAL;
                noDataCondition.CondVal = "___SIN_RESULTADOS___";
                reportChooseFromList.SetConditions(conditions);
                return;
            }

            for (int i = 0; i < reportIds.Count; i++)
            {
                var condition = conditions.Add();
                condition.Alias = "U_SS_IDRPT";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = reportIds[i];

                if (i < reportIds.Count - 1)
                {
                    condition.Relationship = BoConditionRelationship.cr_OR;
                }
            }

            reportChooseFromList.SetConditions(conditions);
        }

        private List<string> GetReportIdsByDepartment(string departmentId)
        {
            var reportIds = new List<string>();
            var sanitizedDepartmentId = departmentId.Replace("'", "''");
            var sql =
                "SELECT DISTINCT T1.\"U_SS_IDRPT\" " +
                "FROM \"@SS_DFRPT_CAB\" T0 " +
                "INNER JOIN \"@SS_DFRPT_DET\" T1 ON T0.\"Code\" = T1.\"Code\" " +
                "WHERE T0.\"U_SS_IDDPT\" = '" + sanitizedDepartmentId + "' " +
                "AND IFNULL(T1.\"U_SS_IDRPT\", '') <> ''";

            Recordset rs = null;
            try
            {
                rs = (Recordset)Globals.rCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    var reportId = (rs.Fields.Item(0).Value ?? string.Empty).ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(reportId))
                    {
                        reportIds.Add(reportId);
                    }

                    rs.MoveNext();
                }
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }
            }

            return reportIds;
        }

        private void HandleChooseFromListSelection(ref ItemEvent pVal)
        {
            var chooseFromListEvent = (IChooseFromListEvent)pVal;
            var selectedObjects = chooseFromListEvent.SelectedObjects;
            if (selectedObjects == null || selectedObjects.Rows.Count == 0)
            {
                return;
            }

            var form = TryGetOpenForm(FormUid);
            if (form == null)
            {
                return;
            }

            if (pVal.ItemUID == DepartmentEditUid)
            {
                var departmentValue = GetFirstAvailableValue(selectedObjects, "Code", "U_Code", "Name");
                ((EditText)form.Items.Item(DepartmentEditUid).Specific).Value = departmentValue;
                ((EditText)form.Items.Item(ReportEditUid).Specific).Value = string.Empty;
                return;
            }

            if (pVal.ItemUID == ReportEditUid)
            {
                var reportValue = GetFirstAvailableValue(selectedObjects, "U_SS_IDRPT", "Code", "Name");
                ((EditText)form.Items.Item(ReportEditUid).Specific).Value = reportValue;
            }
        }

        private static string GetFirstAvailableValue(DataTable table, params string[] columns)
        {
            foreach (var column in columns)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (!string.Equals(table.Columns.Item(i).Name, column, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var value = (table.GetValue(table.Columns.Item(i).Name, 0) ?? string.Empty).ToString();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        return value.Trim();
                    }
                }
            }

            return string.Empty;
        }

        private void OpenConfigurationLoginForm()
        {
            var existing = TryGetOpenForm(LoginFormUid) ?? TryGetOpenForm(ConfigFormUid);

            if (existing != null)
            {
                existing.Visible = true;
                existing.Select();
                return;
            }

            var creationParams = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            creationParams.UniqueID = LoginFormUid;
            creationParams.FormType = LoginFormUid;
            creationParams.BorderStyle = BoFormBorderStyle.fbs_Fixed;

            var form = _app.Forms.AddEx(creationParams);
            form.Title = "Autenticación Configuración";
            form.Width = 340;
            form.Height = 180;

            AddStaticText(form, "lbl_usr", "Usuario", 15, 20, 90);
            AddEditText(form, "txt_usr", 110, 18, 180, false);

            AddStaticText(form, "lbl_pwd", "Clave", 15, 55, 90);
            AddEditText(form, "txt_pwd", 110, 53, 180, true);

            AddButton(form, "btn_ok", "Ingresar", 110, 95, 90);
            AddButton(form, "2", "Cancelar", 210, 95, 80);
            form.Visible = true;
        }

        private void ValidateCredentialsAndOpenConfiguration()
        {
            var loginForm = TryGetOpenForm(LoginFormUid);
            if (loginForm == null)
                return;

            var user = ((EditText)loginForm.Items.Item("txt_usr").Specific).Value;
            var pass = ((EditText)loginForm.Items.Item("txt_pwd").Specific).Value;

            if (user == "admin" && pass == "admin")
            {
                loginForm.Close();
                OpenConfigurationForm();
                _app.StatusBar.SetText("Acceso autorizado.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                return;
            }

            _app.StatusBar.SetText("Credenciales inválidas.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        }

        private void OpenConfigurationForm()
        {
            var existing = TryGetOpenForm(ConfigFormUid);
            if (existing != null)
            {
                existing.Visible = true;
                existing.Select();
                return;
            }

            var creationParams = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            creationParams.UniqueID = ConfigFormUid;
            creationParams.FormType = ConfigFormUid;
            creationParams.BorderStyle = BoFormBorderStyle.fbs_Fixed;

            var form = _app.Forms.AddEx(creationParams);
            form.Title = "Configuración ReportManager";
            form.Width = 460;
            form.Height = 200;

            AddButton(form, "btn_dpts", "Crear tabla ss_dpts y ss_prmtypes", 20, 30, 190);
            AddButton(form, "btn_udos", "Crear estructuras y UDO", 220, 30, 200);

            AddStaticText(form, "lbl_inf", "Ejecute primero ss_dpts y luego estructuras maestras.", 20, 70, 400);
            form.Visible = true;
        }

        private Form TryGetOpenForm(string formUid)
        {
            try
            {
                return _app.Forms.Item(formUid);
            }
            catch
            {
                return null;
            }
        }

        private static void AddStaticText(Form form, string uid, string caption, int left, int top, int width)
        {
            var item = form.Items.Add(uid, BoFormItemTypes.it_STATIC);
            item.Left = left;
            item.Top = top;
            item.Width = width;
            item.Height = 15;
            var text = (StaticText)item.Specific;
            text.Caption = caption;
        }

        private static void AddEditText(Form form, string uid, int left, int top, int width, bool isPassword)
        {
            var item = form.Items.Add(uid, BoFormItemTypes.it_EDIT);
            item.Left = left;
            item.Top = top;
            item.Width = width;
            item.Height = 15;

            var text = (EditText)item.Specific;
            if (isPassword)
            {
                text.IsPassword = true;
            }
        }

        private static void AddButton(Form form, string uid, string caption, int left, int top, int width)
        {
            var item = form.Items.Add(uid, BoFormItemTypes.it_BUTTON);
            item.Left = left;
            item.Top = top;
            item.Width = width;
            item.Height = 19;
            var button = (Button)item.Specific;
            button.Caption = caption;
        }
    }

}
