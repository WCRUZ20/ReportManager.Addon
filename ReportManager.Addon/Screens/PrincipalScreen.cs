using ReportManager.Addon.Core;
using ReportManager.Addon.Core.Embedding;
using ReportManager.Addon.Entidades;
using ReportManager.Addon.Logging;
using ReportManager.Addon.Services;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
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
        private const string ConfigTabUid = "tab_cfg";
        private const string InitTabPaneUid = "1";
        private const string InitialDbUserUid = "txt_dbusr";
        private const string InitialDbPassUid = "txt_dbpwd";
        private const string InitialLoadSapUid = "chk_ldsap";
        private const string SaveConfigButtonUid = "btn_save";
        private const string ConfigFormUid = "RM_CFG_FORM";
        private const string DepartmentEditUid = "edt_dpt";
        private const string ReportEditUid = "edt_rpt";
        private const string DepartmentChooseFromListUid = "RM_CFL_DPT";
        private const string ReportChooseFromListUid = "RM_CFL_RPT";
        private readonly Application _app;
        private readonly Logger _log;
        private readonly PrincipalFormController _principalFormController;
        private readonly ConfigurationMetadataService _configurationMetadataService;
        private readonly ReportParameterMapper _reportParameterMapper;
        private readonly SAPbobsCOM.Company _company;
        private readonly SrfFormLoader _srfFormLoader;

        private const string DepartmentComboUid = "cmb_dpt";
        private const string ReportsGridUid = "grd_rpt";
        private const string ReportsDataTableUid = "DT_RPTS";
        private const string EmbeddedBoxUid = "bx_rptdtl";
        private const string EmbeddedTitleUid = "lbl_rptttl";
        private const string EmbeddedIdLabelUid = "lbl_rptid";
        private const string EmbeddedNameLabelUid = "lbl_rptnam";
        private const string EmbeddedIdValueUid = "txt_rptid";
        private const string EmbeddedNameValueUid = "txt_rptnam";
        private const string ParentFormDataSourceUid = "UD_PARENT";

        private Form1 formPrueba;
        private readonly EmbeddedWinFormRegistry _embeddedRegistry = new EmbeddedWinFormRegistry();
        private const string EmbeddedHostFormUid = "RM_EMB_RPT";
        private const string EmbeddedHostTitle = "Generación de Reporte";

        public PrincipalScreen(
            Application app,
            Logger log,
            PrincipalFormController principalFormController,
            ConfigurationMetadataService configurationMetadataService,
            ReportParameterMapper reportParameterMapper,
            SAPbobsCOM.Company company,
            SrfFormLoader srfFormLoader)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _principalFormController = principalFormController ?? throw new ArgumentNullException(nameof(principalFormController));
            _configurationMetadataService = configurationMetadataService ?? throw new ArgumentNullException(nameof(configurationMetadataService));
            _reportParameterMapper = reportParameterMapper ?? throw new ArgumentNullException(nameof(reportParameterMapper));
            _company = company ?? throw new ArgumentNullException(nameof(company));
            _srfFormLoader = srfFormLoader ?? throw new ArgumentNullException(nameof(srfFormLoader));

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
                if (pVal.BeforeAction
                    && pVal.EventType == BoEventTypes.et_FORM_CLOSE
                    && HasOpenChildForms(formUID))
                {
                    bubbleEvent = false;
                    _app.StatusBar.SetText("No puede cerrar este formulario mientras existan formularios hijos abiertos.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                if (formUID == FormUid
                    && pVal.EventType == BoEventTypes.et_FORM_VISIBLE
                    && !pVal.BeforeAction)
                {
                    var form = TryGetOpenForm(FormUid);
                    var ItemsCount = form.Items.Count; //abre algo como un form vacio al comienzo, despues un segundo form con los elementos definidos en el xml
                    if (form != null && ItemsCount != 0) //por eso valido aqui cuando el # elementos es diferente de 0
                    {
                        LoadDepartmentsCombo(form);
                        LoadReportsGrid(form);
                    }
                    return;
                }

                if (formUID == FormUid
                    && pVal.EventType == BoEventTypes.et_COMBO_SELECT
                    && pVal.ItemUID == DepartmentComboUid
                    && pVal.ActionSuccess)
                {
                    var form = TryGetOpenForm(FormUid);
                    if (form != null)
                    {
                        LoadReportsGrid(form);
                    }
                    return;
                }

                if (_reportParameterMapper.IsMappingForm(formUID)
                    && pVal.EventType == BoEventTypes.et_ITEM_PRESSED
                    && pVal.ActionSuccess
                    && _reportParameterMapper.IsParameterButton(pVal.ItemUID))
                {
                    _reportParameterMapper.OpenQuerySelector(formUID, pVal.ItemUID);
                    return;
                }

                if (_reportParameterMapper.IsMappingForm(formUID)
                    && pVal.EventType == BoEventTypes.et_ITEM_PRESSED
                    && pVal.ActionSuccess
                    && _reportParameterMapper.IsGenerateReportButton(pVal.ItemUID))
                {
                    _reportParameterMapper.GenerateSelectedReport(formUID);

                    //Form Incrustado
                    //try
                    //{
                    //    var host = _embeddedRegistry.GetOrCreate(_app, _log, EmbeddedHostFormUid, EmbeddedHostTitle);
                    //    host.ShowOrFocus(() => new ReportManager.Addon.Form1(), width: 900, height: 650);
                    //}
                    //catch (Exception ex)
                    //{
                    //    _log.Error("No se pudo incrustar Form1.", ex);
                    //    _app.StatusBar.SetText("No se pudo incrustar el formulario de reporte: " + ex.Message,
                    //        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    //}

                    return;
                }

                if (formUID == EmbeddedHostFormUid
                    && pVal.EventType == BoEventTypes.et_FORM_CLOSE
                    && pVal.BeforeAction)
                {
                    _embeddedRegistry.DisposeHost(EmbeddedHostFormUid);
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

                if (formUID == FormUid
                    && !pVal.BeforeAction
                    && (pVal.EventType == BoEventTypes.et_CLICK || pVal.EventType == BoEventTypes.et_DOUBLE_CLICK)
                    && pVal.ItemUID == ReportsGridUid
                    && pVal.Row >= 0)
                {
                    var form = TryGetOpenForm(FormUid);
                    if (form != null)
                    {
                        //ShowEmbeddedReportForm(form, pVal.Row);
                        _reportParameterMapper.ShowFromSelectedReportRow(form, ReportsGridUid, pVal.Row);
                    }

                    return;
                }

                if (_reportParameterMapper.IsQueryPickerForm(formUID)
                    && !pVal.BeforeAction
                    && (pVal.EventType == BoEventTypes.et_CLICK || pVal.EventType == BoEventTypes.et_DOUBLE_CLICK)
                    && _reportParameterMapper.IsQueryPickerGrid(pVal.ItemUID))
                {
                    if (pVal.EventType == BoEventTypes.et_CLICK && !string.IsNullOrWhiteSpace(pVal.ColUID) && pVal.Row == -1)
                    {
                        _reportParameterMapper.UpdateQueryPickerSelectedColumn(formUID, pVal.ColUID);
                    }

                    if (pVal.EventType == BoEventTypes.et_DOUBLE_CLICK && pVal.Row >= 0)
                    {
                        _reportParameterMapper.ApplyQuerySelection(formUID, pVal.Row);
                    }

                    return;
                }

                if (_reportParameterMapper.IsQueryPickerForm(formUID)
                    && !pVal.BeforeAction
                    && pVal.EventType == BoEventTypes.et_KEY_DOWN
                    && _reportParameterMapper.IsQueryPickerSearchItem(pVal.ItemUID))
                {
                    _reportParameterMapper.RefreshQueryPickerGrid(formUID);
                    return;
                }

                if (_reportParameterMapper.IsQueryPickerForm(formUID)
                    && pVal.EventType == BoEventTypes.et_VALIDATE
                    && pVal.ActionSuccess
                    && _reportParameterMapper.IsQueryPickerSearchItem(pVal.ItemUID))
                {
                    _reportParameterMapper.RefreshQueryPickerGrid(formUID);
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
                        _configurationMetadataService.CreateGeneralConfigurationTable();
                        _configurationMetadataService.CreateParamterTypeValuesTable();
                        return;
                    }

                    if (pVal.ItemUID == "btn_udos")
                    {
                        _configurationMetadataService.CreateReportConfigurationStructures();
                        return;
                    }

                    if (pVal.ItemUID == SaveConfigButtonUid)
                    {
                        SaveGeneralConfiguration();
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

        private bool HasOpenChildForms(string parentFormUid)
        {
            if (string.IsNullOrWhiteSpace(parentFormUid))
            {
                return false;
            }

            for (int i = 0; i < _app.Forms.Count; i++)
            {
                Form form = null;
                try
                {
                    form = _app.Forms.Item(i);
                }
                catch
                {
                    continue;
                }

                if (form == null || string.Equals(form.UniqueID, parentFormUid, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    var parentUid = form.DataSources.UserDataSources.Item(ParentFormDataSourceUid).ValueEx;
                    if (string.Equals(parentUid, parentFormUid, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        private void LoadDepartmentsCombo(Form form)
        {
            SAPbobsCOM.Recordset recordset = null;

            try
            {
                var combo = (ComboBox)form.Items.Item(DepartmentComboUid).Specific;
                var comboItem = form.Items.Item(DepartmentComboUid);
                comboItem.DisplayDesc = true;

                while (combo.ValidValues.Count > 0)
                {
                    combo.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                string query = "";

                if (_company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    query = "SELECT \"Code\", \"Name\" FROM \"@SS_DPTS\" ORDER BY \"Code\"";
                }
                else
                {
                    query = "SELECT Code, Name FROM [@SS_DPTS] ORDER BY Code";
                }
                recordset = (SAPbobsCOM.Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery(query);

                while (!recordset.EoF)
                {
                    var code = Convert.ToString(recordset.Fields.Item("Code").Value);
                    var name = Convert.ToString(recordset.Fields.Item("Name").Value);

                    if (!string.IsNullOrWhiteSpace(code))
                    {
                        combo.ValidValues.Add(code, name ?? string.Empty);
                    }

                    recordset.MoveNext();
                }

                if (combo.ValidValues.Count > 0)
                {
                    combo.Select(0, BoSearchKey.psk_Index);
                }
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo cargar el combo de departamentos.", ex);
                _app.StatusBar.SetText("No se pudo cargar SS_DPTS en el combo.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (recordset != null)
                {
                    Marshal.ReleaseComObject(recordset);
                }
            }
        }

        private void ShowEmbeddedReportForm(Form form, int row)
        {
            var grid = (Grid)form.Items.Item(ReportsGridUid).Specific;
            if (grid.DataTable == null || grid.DataTable.Rows.Count <= row)
            {
                return;
            }

            EnsureEmbeddedReportControls(form);

            var reportId = Convert.ToString(grid.DataTable.GetValue("U_SS_IDRPT", row)) ?? string.Empty;
            var reportName = Convert.ToString(grid.DataTable.GetValue("U_SS_NOMBRPT", row)) ?? string.Empty;

            ((EditText)form.Items.Item(EmbeddedIdValueUid).Specific).Value = reportId;
            ((EditText)form.Items.Item(EmbeddedNameValueUid).Specific).Value = reportName;
        }

        private void ClearEmbeddedReportData(Form form)
        {
            if (!HasItem(form, EmbeddedIdValueUid) || !HasItem(form, EmbeddedNameValueUid))
            {
                return;
            }

            ((EditText)form.Items.Item(EmbeddedIdValueUid).Specific).Value = string.Empty;
            ((EditText)form.Items.Item(EmbeddedNameValueUid).Specific).Value = string.Empty;
        }

        private static void EnsureEmbeddedReportControls(Form form)
        {
            if (!HasItem(form, EmbeddedBoxUid))
            {
                var panel = form.Items.Add(EmbeddedBoxUid, BoFormItemTypes.it_RECTANGLE);
                panel.Left = 337;
                panel.Top = 66;
                panel.Width = 320;
                panel.Height = 120;
            }

            if (!HasItem(form, EmbeddedTitleUid))
            {
                AddStaticText(form, EmbeddedTitleUid, "Detalle del reporte", 347, 76, 180);
            }

            if (!HasItem(form, EmbeddedIdLabelUid))
            {
                AddStaticText(form, EmbeddedIdLabelUid, "Id:", 347, 100, 30);
            }

            if (!HasItem(form, EmbeddedIdValueUid))
            {
                AddEditText(form, EmbeddedIdValueUid, 385, 98, 250, false);
                form.Items.Item(EmbeddedIdValueUid).Enabled = false;
            }

            if (!HasItem(form, EmbeddedNameLabelUid))
            {
                AddStaticText(form, EmbeddedNameLabelUid, "Nombre:", 347, 124, 45);
            }

            if (!HasItem(form, EmbeddedNameValueUid))
            {
                AddEditText(form, EmbeddedNameValueUid, 395, 122, 240, false);
                form.Items.Item(EmbeddedNameValueUid).Enabled = false;
            }
        }

        private static bool HasItem(Form form, string itemUid)
        {
            try
            {
                form.Items.Item(itemUid);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void LoadReportsGrid(Form form)
        {
            try
            {
                var combo = (ComboBox)form.Items.Item(DepartmentComboUid).Specific;
                var selectedDepartment = combo.Selected == null ? string.Empty : combo.Selected.Value;

                if (string.IsNullOrWhiteSpace(selectedDepartment))
                {
                    return;
                }

                var grid = (Grid)form.Items.Item(ReportsGridUid).Specific;
                var dataTable = GetOrCreateDataTable(form, ReportsDataTableUid);

                var escapedDepartment = selectedDepartment.Replace("'", "''");
                var query = BuildReportsByDepartmentQuery(escapedDepartment);
                dataTable.ExecuteQuery(query);

                grid.DataTable = dataTable;
                grid.SelectionMode = BoMatrixSelect.ms_Single;
                grid.Item.Enabled = false;

                if (grid.Columns.Count > 0)
                {
                    grid.Columns.Item("U_SS_IDRPT").TitleObject.Caption = "Id Reporte";
                }

                if (grid.Columns.Count > 1)
                {
                    grid.Columns.Item("U_SS_NOMBRPT").TitleObject.Caption = "Nombre Reporte";
                }

                grid.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo cargar el grid de reportes.", ex);
                _app.StatusBar.SetText("No se pudo cargar la grilla de reportes.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private string BuildReportsByDepartmentQuery(string departmentCode)
        {
            if (_company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                return "select T2.\"U_SS_IDRPT\", T2.\"U_SS_NOMBRPT\"from \"@SS_PRM_CAB\" T2 where T2.\"U_SS_IDDPT\" = 'departmentCode' order by 1";
            }

            return $@"select T2.U_SS_IDRPT, T2.U_SS_NOMBRPT
                from [@SS_PRM_CAB] T2 
                where T2.U_SS_IDDPT = '{departmentCode}'
                order by 1";
        }

        private static DataTable GetOrCreateDataTable(Form form, string dataTableUid)
        {
            try
            {
                return form.DataSources.DataTables.Item(dataTableUid);
            }
            catch
            {
                return form.DataSources.DataTables.Add(dataTableUid);
            }
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

            OpenFormFromSrf(LoginFormUid, "ConfigurationLogin.srf");
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
            OpenFormFromSrf(ConfigFormUid, "Configuration.srf");
        }

        private void SaveGeneralConfiguration()
        {
            var configForm = TryGetOpenForm(ConfigFormUid);
            if (configForm == null)
            {
                return;
            }

            var dbUser = ((EditText)configForm.Items.Item(InitialDbUserUid).Specific).Value;
            var dbPwd = ((EditText)configForm.Items.Item(InitialDbPassUid).Specific).Value;
            var loadSapCheckbox = (CheckBox)configForm.Items.Item(InitialLoadSapUid).Specific;
            var loadSap = loadSapCheckbox.Checked ? loadSapCheckbox.ValOn : loadSapCheckbox.ValOff;

            try
            {
                _configurationMetadataService.CreateGeneralConfigurationTable(); //se peude validar si la tabla existe, ya no haga esta accion

                UpsertConfigurationValue(InitialDbUserUid, dbUser);
                Globals.dbuser = dbUser;
                UpsertConfigurationValue(InitialDbPassUid, dbPwd);
                Globals.pwduser = dbPwd;
                UpsertConfigurationValue(InitialLoadSapUid, loadSap);
                Globals.chkrptSAP = loadSap;

                _app.StatusBar.SetText("Configuración guardada en SS_CONFG_RM.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo guardar la configuración general.", ex);
                _app.StatusBar.SetText("No se pudo guardar la configuración.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void UpsertConfigurationValue(string key, string value)
        {
            var escapedKey = EscapeSqlValue(key);
            var escapedValue = EscapeSqlValue(value ?? string.Empty);

            var query = _company.DbServerType == BoDataServerTypes.dst_HANADB
                ? $@"UPSERT ""@SS_CONFG_RM"" (""Code"", ""Name"", ""U_SS_CLAVE"", ""U_SS_VALOR"")
                    VALUES ('{escapedKey}', '{escapedKey}', '{escapedKey}', '{escapedValue}') WITH PRIMARY KEY"
                                    : $@"MERGE [@SS_CONFG_RM] AS target
                    USING (SELECT '{escapedKey}' AS CodeValue, '{escapedValue}' AS DataValue) AS source
                    ON target.[Code] = source.CodeValue
                    WHEN MATCHED THEN
                        UPDATE SET
                            target.[Name] = source.CodeValue,
                            target.[U_SS_CLAVE] = source.CodeValue,
                            target.[U_SS_VALOR] = source.DataValue
                    WHEN NOT MATCHED THEN
                        INSERT ([Code], [Name], [U_SS_CLAVE], [U_SS_VALOR])
                        VALUES (source.CodeValue, source.CodeValue, source.CodeValue, source.DataValue);";

            SAPbobsCOM.Recordset recordset = null;

            try
            {
                recordset = (SAPbobsCOM.Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery(query);
            }
            finally
            {
                if (recordset != null)
                {
                    Marshal.ReleaseComObject(recordset);
                }
            }
        }

        private static string EscapeSqlValue(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }


        private void OpenFormFromSrf(string formUid, string srfName)
        {
            var existing = TryGetOpenForm(formUid);
            if (existing != null)
            {
                existing.Visible = true;
                existing.Select();
                return;
            }

            var srfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Forms", srfName);
            _srfFormLoader.LoadFromFile(srfPath);

            var loadedForm = TryGetOpenForm(formUid);
            if (loadedForm != null)
            {
                loadedForm.Visible = true;
                loadedForm.Select();
            }
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
