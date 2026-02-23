using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using ReportManager.Addon.Forms;
using ReportManager.Addon.Logging;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Services
{
    public sealed class ReportParameterMapper
    {
        private const string ParametersPrefix = "prm_";
        private const string MappingFormType = "RM_PRM_MAP";
        private const string MappingHeaderUid = "lbl_hdr";
        private const string QueryResultFormType = "RM_QRY_PICK";
        private const string QueryResultGridUid = "grd_qry";
        private const string QueryResultDataTableUid = "DT_QRY";
        private const string QueryResultSourceDataSource = "UD_SRC";
        private const string ParentFormDataSourceUid = "UD_PARENT";
        private const string QuerySearchLabelUid = "lbl_qsrch";
        private const string QuerySearchEditUid = "edt_qsrch";
        private const string GenerateReportButtonUid = "btn_exerpt";
        private const string GenerateReportButtonCaption = "Generar reporte";
        private const string ReportsBasePath = @"C:\Reportes SAP";
        private const int ParameterRowHeight = 24;
        private const int MappingFormMinHeight = 180;
        private const int MappingFormBottomPadding = 72;
        private ReportDocument reportDocument;

        private readonly Application _app;
        private readonly Logger _log;
        private readonly SAPbobsCOM.Company _company;
        private readonly Dictionary<string, ParameterUiContext> _parameterContexts = new Dictionary<string, ParameterUiContext>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, QueryPickerContext> _queryPickerContexts = new Dictionary<string, QueryPickerContext>(StringComparer.OrdinalIgnoreCase);

        private string _mappingFormUid;

        public ReportParameterMapper(Application app, Logger log, SAPbobsCOM.Company company)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _company = company ?? throw new ArgumentNullException(nameof(company));
        }

        public bool IsParameterButton(string itemUid)
        {
            return !string.IsNullOrWhiteSpace(itemUid)
                && itemUid.StartsWith(ParametersPrefix + "btn_", StringComparison.OrdinalIgnoreCase);
        }

        public bool IsQueryPickerForm(string formUid)
        {
            return !string.IsNullOrWhiteSpace(formUid)
                && formUid.StartsWith(QueryResultFormType, StringComparison.OrdinalIgnoreCase);
        }

        public bool IsQueryPickerGrid(string itemUid)
        {
            return string.Equals(itemUid, QueryResultGridUid, StringComparison.OrdinalIgnoreCase);
        }

        public bool IsQueryPickerSearchItem(string itemUid)
        {
            return string.Equals(itemUid, QuerySearchEditUid, StringComparison.OrdinalIgnoreCase);
        }

        public bool IsMappingForm(string formUid)
        {
            return !string.IsNullOrWhiteSpace(formUid)
                && formUid.StartsWith(MappingFormType, StringComparison.OrdinalIgnoreCase);
        }

        public bool IsGenerateReportButton(string itemUid)
        {
            return string.Equals(itemUid, GenerateReportButtonUid, StringComparison.OrdinalIgnoreCase);
        }

        public void GenerateSelectedReport(string mappingFormUid)
        {
            try
            {
                var form = _app.Forms.Item(mappingFormUid);
                var reportInfo = GetCurrentReportInfo(form);
                if (reportInfo == null)
                {
                    _app.StatusBar.SetText("No se pudo identificar el reporte seleccionado.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }

                var reportFilePath = ResolveReportFilePath(reportInfo.ReportCode, reportInfo.ReportName);
                if (string.IsNullOrWhiteSpace(reportFilePath))
                {
                    _app.StatusBar.SetText("No se encontró el archivo de Crystal Report para el reporte seleccionado.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }

                var parameterDefinitions = GetReportParameters(reportInfo.ReportCode);
                var parameterValues = BuildParameterValues(form, parameterDefinitions);

                reportDocument = new ReportDocument();
                reportDocument.Load(reportFilePath);
                //ApplyParameters(reportDocument, parameterValues);

                var viewerForm = new CrystalReportViewerForm(reportDocument);
                viewerForm.Show();

                _app.StatusBar.SetText("Reporte abierto correctamente.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo abrir el Crystal Report seleccionado.", ex);
                _app.StatusBar.SetText("No se pudo abrir el reporte seleccionado.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void OpenOrRefreshMappingForm(string parentFormUid, string reportCode, string reportName, List<ReportParameterDefinition> parameters)
        {
            CloseMappingFormIfOpen();

            var uid = MappingFormType + DateTime.Now.Ticks;
            var creationParams = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            creationParams.UniqueID = uid;
            creationParams.FormType = MappingFormType;
            creationParams.BorderStyle = BoFormBorderStyle.fbs_Fixed;

            var form = _app.Forms.AddEx(creationParams);
            SetParentForm(form, parentFormUid);
            form.Title = "Parámetros de reporte";
            form.Left = 680;
            form.Top = 90;
            form.Width = 550;

            var headerItem = form.Items.Add(MappingHeaderUid, BoFormItemTypes.it_STATIC);
            headerItem.Left = 12;
            headerItem.Top = 12;
            headerItem.Width = 600;
            ((StaticText)headerItem.Specific).Caption = $"Reporte: {reportCode} - {reportName}";

            RenderParameters(form, parameters);
            AddGenerateReportButton(form, parameters.Count);
            ResizeMappingFormHeight(form, parameters.Count);

            form.Visible = true;
            _mappingFormUid = uid;
        }

        private Dictionary<string, object> BuildParameterValues(Form mappingForm, List<ReportParameterDefinition> definitions)
        {
            var values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (mappingForm == null || definitions == null || definitions.Count == 0)
            {
                return values;
            }

            for (int i = 0; i < definitions.Count; i++)
            {
                var definition = definitions[i];
                if (string.IsNullOrWhiteSpace(definition?.ParamId))
                {
                    continue;
                }

                var valueUid = ParametersPrefix + "val_" + i.ToString("00");
                if (!HasItem(mappingForm, valueUid))
                {
                    continue;
                }

                var parameterValue = GetUiValue(mappingForm, valueUid, definition.Type);
                if (parameterValue != null)
                {
                    values[definition.ParamId] = parameterValue;
                }
            }

            return values;
        }

        private object GetUiValue(Form mappingForm, string valueUid, string parameterType)
        {
            if (IsBooleanType(parameterType))
            {
                return ((CheckBox)mappingForm.Items.Item(valueUid).Specific).Checked;
            }

            var rawValue = ((EditText)mappingForm.Items.Item(valueUid).Specific).Value;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return null;
            }

            if (string.Equals(parameterType, "DATE", StringComparison.OrdinalIgnoreCase)
                && DateTime.TryParseExact(rawValue, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out var dateValue))
            {
                return dateValue;
            }

            if (IsNumericType(parameterType) && decimal.TryParse(rawValue, out var numericValue))
            {
                return numericValue;
            }

            return rawValue;
        }

        private static void ApplyParameters(ReportDocument reportDocument, Dictionary<string, object> parameterValues)
        {
            if (reportDocument == null || parameterValues == null || parameterValues.Count == 0)
            {
                return;
            }

            foreach (ParameterFieldDefinition parameter in reportDocument.DataDefinition.ParameterFields)
            {
                if (!parameterValues.TryGetValue(parameter.Name, out var value))
                {
                    continue;
                }

                var discreteValue = new ParameterDiscreteValue { Value = value };
                var currentValues = new ParameterValues();
                currentValues.Add(discreteValue);
                parameter.ApplyCurrentValues(currentValues);
            }
        }

        public void ShowFromSelectedReportRow(Form principalForm, string reportsGridUid, int selectedRow)
        {
            if (principalForm == null)
            {
                return;
            }

            try
            {
                var grid = (Grid)principalForm.Items.Item(reportsGridUid).Specific;
                var reportCode = Convert.ToString(grid.DataTable.GetValue("U_SS_IDRPT", selectedRow));
                var reportName = Convert.ToString(grid.DataTable.GetValue("U_SS_NOMBRPT", selectedRow));
                if (string.IsNullOrWhiteSpace(reportCode))
                {
                    return;
                }

                var parameters = GetReportParameters(reportCode);
                OpenOrRefreshMappingForm(principalForm.UniqueID, reportCode, reportName, parameters);
            }
            catch (Exception ex)
            {
                _log.Error("No se pudieron mapear los parámetros del reporte.", ex);
                _app.StatusBar.SetText("No se pudieron cargar los parámetros del reporte.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public void CloseMappingFormIfOpen()
        {
            if (string.IsNullOrWhiteSpace(_mappingFormUid))
            {
                return;
            }

            try
            {
                var form = _app.Forms.Item(_mappingFormUid);
                form.Close();
            }
            catch
            {
            }

            _mappingFormUid = null;
            _parameterContexts.Clear();
        }

        public void OpenQuerySelector(string sourceFormUid, string buttonItemUid)
        {
            if (!_parameterContexts.TryGetValue(buttonItemUid, out var context) || string.IsNullOrWhiteSpace(context.Query))
            {
                return;
            }

            var queryFormUid = QueryResultFormType + DateTime.Now.Ticks;
            var creationParams = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            creationParams.UniqueID = queryFormUid;
            creationParams.FormType = QueryResultFormType;
            creationParams.BorderStyle = BoFormBorderStyle.fbs_Sizable;

            var form = _app.Forms.AddEx(creationParams);
            SetParentForm(form, sourceFormUid);
            form.Title = "Seleccionar valor";
            form.Width = 550;
            form.Height = 400;

            var searchLabelItem = form.Items.Add(QuerySearchLabelUid, BoFormItemTypes.it_STATIC);
            searchLabelItem.Left = 10;
            searchLabelItem.Top = 10;
            searchLabelItem.Width = 500;
            ((StaticText)searchLabelItem.Specific).Caption = "Buscar por: Columna 1";

            var searchItem = form.Items.Add(QuerySearchEditUid, BoFormItemTypes.it_EDIT);
            searchItem.Left = 10;
            searchItem.Top = 26;
            searchItem.Width = 500;
            searchItem.Height = 15;

            var gridItem = form.Items.Add(QueryResultGridUid, BoFormItemTypes.it_GRID);
            gridItem.Left = 10;
            gridItem.Top = 46;
            gridItem.Width = 500;
            gridItem.Height = 294;

            var grid = (Grid)gridItem.Specific;
            var dt = form.DataSources.DataTables.Add(QueryResultDataTableUid);
            dt.ExecuteQuery(context.Query);
            grid.DataTable = dt;
            grid.SelectionMode = BoMatrixSelect.ms_Single;
            grid.Item.Enabled = false;
            grid.AutoResizeColumns();
            string defaultColumn = null;
            if (grid.DataTable != null && grid.DataTable.Columns.Count > 0)
            {
                defaultColumn = grid.DataTable.Columns.Item(0).Name;
            }

            _queryPickerContexts[queryFormUid] = new QueryPickerContext
            {
                BaseQuery = context.Query,
                SelectedColumn = defaultColumn,
            };

            UpdateQueryPickerSearchCaption(form, defaultColumn);

            var source = sourceFormUid + "|" + context.ValueItemUid;
            form.DataSources.UserDataSources.Add(QueryResultSourceDataSource, BoDataType.dt_LONG_TEXT, 200);
            form.DataSources.UserDataSources.Item(QueryResultSourceDataSource).ValueEx = source;
            form.Visible = true;
        }

        private static void SetParentForm(Form form, string parentFormUid)
        {
            if (form == null || string.IsNullOrWhiteSpace(parentFormUid))
            {
                return;
            }

            try
            {
                form.DataSources.UserDataSources.Add(ParentFormDataSourceUid, BoDataType.dt_SHORT_TEXT, 100);
            }
            catch
            {
            }

            try
            {
                form.DataSources.UserDataSources.Item(ParentFormDataSourceUid).ValueEx = parentFormUid;
            }
            catch
            {
            }
        }

        public void UpdateQueryPickerSelectedColumn(string queryFormUid, string columnUid)
        {
            if (string.IsNullOrWhiteSpace(queryFormUid) || string.IsNullOrWhiteSpace(columnUid))
            {
                return;
            }

            if (!_queryPickerContexts.TryGetValue(queryFormUid, out var context))
            {
                return;
            }

            context.SelectedColumn = columnUid;
            _queryPickerContexts[queryFormUid] = context;

            try
            {
                var form = _app.Forms.Item(queryFormUid);
                UpdateQueryPickerSearchCaption(form, columnUid);
                RefreshQueryPickerGrid(queryFormUid);
            }
            catch
            {
            }
        }

        public void RefreshQueryPickerGrid(string queryFormUid)
        {
            if (!_queryPickerContexts.TryGetValue(queryFormUid, out var context))
            {
                return;
            }

            try
            {
                var form = _app.Forms.Item(queryFormUid);
                var searchValue = string.Empty;
                if (HasItem(form, QuerySearchEditUid))
                {
                    searchValue = ((EditText)form.Items.Item(QuerySearchEditUid).Specific).Value;
                }

                var grid = (Grid)form.Items.Item(QueryResultGridUid).Specific;
                var dt = form.DataSources.DataTables.Item(QueryResultDataTableUid);
                var query = BuildFilteredQuery(context.BaseQuery, context.SelectedColumn, searchValue);
                dt.ExecuteQuery(query);
                grid.DataTable = dt;
                grid.AutoResizeColumns();
                UpdateQueryPickerSearchCaption(form, context.SelectedColumn);
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo filtrar la grilla de consultas.", ex);
            }
        }

        private string BuildFilteredQuery(string baseQuery, string selectedColumn, string searchValue)
        {
            if (string.IsNullOrWhiteSpace(baseQuery)
                || string.IsNullOrWhiteSpace(selectedColumn)
                || string.IsNullOrWhiteSpace(searchValue))
            {
                return baseQuery;
            }

            var escapedValue = searchValue.Replace("'", "''");

            if (_company.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                var escapedColumn = selectedColumn.Replace("\"", "\"\"");
                return $"select * from ({baseQuery}) T where lower(to_nvarchar(T.\"{escapedColumn}\")) like '%{escapedValue.ToLowerInvariant()}%'";
            }

            var escapedSqlColumn = selectedColumn.Replace("]", "]]");
            return $"select * from ({baseQuery}) T where lower(cast(T.[{escapedSqlColumn}] as nvarchar(max))) like '%{escapedValue.ToLowerInvariant()}%'";
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

        private void UpdateQueryPickerSearchCaption(Form form, string selectedColumn)
        {
            if (!HasItem(form, QuerySearchLabelUid))
            {
                return;
            }

            var caption = "Buscar por: Columna 1";
            try
            {
                var grid = (Grid)form.Items.Item(QueryResultGridUid).Specific;
                if (!string.IsNullOrWhiteSpace(selectedColumn) && grid.DataTable != null)
                {
                    var index = GetColumnIndex(grid.DataTable, selectedColumn);
                    if (index >= 0)
                    {
                        caption = "Buscar por: " + grid.Columns.Item(index).TitleObject.Caption;
                    }
                }
            }
            catch
            {
            }

            ((StaticText)form.Items.Item(QuerySearchLabelUid).Specific).Caption = caption;
        }

        private static int GetColumnIndex(DataTable dataTable, string columnName)
        {
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                if (string.Equals(dataTable.Columns.Item(i).Name, columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        public void ApplyQuerySelection(string queryFormUid, int row)
        {
            try
            {
                var queryForm = _app.Forms.Item(queryFormUid);
                var source = queryForm.DataSources.UserDataSources.Item(QueryResultSourceDataSource).ValueEx;
                var parts = source.Split('|');
                if (parts.Length != 2)
                {
                    return;
                }

                var sourceForm = _app.Forms.Item(parts[0]);
                var valueUid = parts[1];
                var queryGrid = (Grid)queryForm.Items.Item(QueryResultGridUid).Specific;
                if (queryGrid.DataTable.Columns.Count == 0)
                {
                    return;
                }

                var selectedValue = Convert.ToString(queryGrid.DataTable.GetValue(0, row));
                if (_parameterContexts.TryGetValue(valueUid, out var valueContext)
                    && valueContext.ControlType == ParameterControlType.CheckBox)
                {
                    var checkBox = (CheckBox)sourceForm.Items.Item(valueUid).Specific;
                    checkBox.Checked = string.Equals(selectedValue, "Y", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(selectedValue, "1", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(selectedValue, "TRUE", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(selectedValue, "T", StringComparison.OrdinalIgnoreCase);
                }
                else
                {
                    ((EditText)sourceForm.Items.Item(valueUid).Specific).Value = selectedValue;
                }

                if (_parameterContexts.TryGetValue(valueUid, out var context)
                    && !string.IsNullOrWhiteSpace(context.DescriptionItemUid)
                    && !string.IsNullOrWhiteSpace(context.DescriptionQuery))
                {
                    var description = ExecuteScalar(ReplaceFiltro(context.DescriptionQuery, selectedValue));
                    ((EditText)sourceForm.Items.Item(context.DescriptionItemUid).Specific).Value = description;
                }

                _queryPickerContexts.Remove(queryFormUid);
                queryForm.Close();
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo aplicar la selección de consulta.", ex);
            }
        }

        private static bool IsBooleanType(string parameterType)
        {
            return string.Equals(parameterType, "BOOL", StringComparison.OrdinalIgnoreCase)
                || string.Equals(parameterType, "BOOLEAN", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsNumericType(string parameterType)
        {
            return string.Equals(parameterType, "NUMERIC", StringComparison.OrdinalIgnoreCase)
                || string.Equals(parameterType, "NUMBER", StringComparison.OrdinalIgnoreCase)
                || string.Equals(parameterType, "INT", StringComparison.OrdinalIgnoreCase)
                || string.Equals(parameterType, "INTEGER", StringComparison.OrdinalIgnoreCase)
                || string.Equals(parameterType, "DECIMAL", StringComparison.OrdinalIgnoreCase);
        }

        private List<ReportParameterDefinition> GetReportParameters(string reportCode)
        {
            var result = new List<ReportParameterDefinition>();
            Recordset recordset = null;

            try
            {
                var escapedReportCode = reportCode.Replace("'", "''");
                var query = _company.DbServerType == BoDataServerTypes.dst_HANADB
                    ? "select \"LineId\", \"U_SS_IDPARAM\", \"U_SS_DSCPARAM\", \"U_SS_TIPO\", \"U_SS_QUERY\", \"U_SS_DESC\", \"U_SS_QUERYD\", \"U_SS_ACTIVO\" from \"@SS_PRM_DET\" where \"Code\" = '{escapedReportCode}' order by \"LineId\""
                    : $@"select LineId, U_SS_IDPARAM, U_SS_DSCPARAM, U_SS_TIPO, U_SS_QUERY, U_SS_DESC, U_SS_QUERYD, U_SS_ACTIVO from [@SS_PRM_DET] where Code = '{escapedReportCode}' order by LineId";

                recordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery(query);

                while (!recordset.EoF)
                {
                    var isActive = Convert.ToString(recordset.Fields.Item("U_SS_ACTIVO").Value);
                    if (!string.Equals(isActive, "N", StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(new ReportParameterDefinition
                        {
                            ParamId = Convert.ToString(recordset.Fields.Item("U_SS_IDPARAM").Value),
                            Description = Convert.ToString(recordset.Fields.Item("U_SS_DSCPARAM").Value),
                            Type = Convert.ToString(recordset.Fields.Item("U_SS_TIPO").Value),
                            Query = Convert.ToString(recordset.Fields.Item("U_SS_QUERY").Value),
                            ShowDescription = string.Equals(Convert.ToString(recordset.Fields.Item("U_SS_DESC").Value), "Y", StringComparison.OrdinalIgnoreCase),
                            DescriptionQuery = Convert.ToString(recordset.Fields.Item("U_SS_QUERYD").Value),
                        });
                    }

                    recordset.MoveNext();
                }
            }
            finally
            {
                if (recordset != null)
                {
                    Marshal.ReleaseComObject(recordset);
                }
            }

            return result;
        }

        private void RenderParameters(Form form, List<ReportParameterDefinition> parameters)
        {
            _parameterContexts.Clear();
            var baseLeft = 12;
            var nextTop = 38;

            for (int i = 0; i < parameters.Count; i++)
            {
                var prm = parameters[i];
                var suffix = i.ToString("00");
                var lblUid = ParametersPrefix + "lbl_" + suffix;
                var valUid = ParametersPrefix + "val_" + suffix;
                var btnUid = ParametersPrefix + "btn_" + suffix;
                var descUid = ParametersPrefix + "dsc_" + suffix;

                var labelItem = form.Items.Add(lblUid, BoFormItemTypes.it_STATIC);
                labelItem.Left = baseLeft;
                labelItem.Top = nextTop + 2;
                labelItem.Width = 120;
                ((StaticText)labelItem.Specific).Caption = string.IsNullOrWhiteSpace(prm.Description) ? prm.ParamId : prm.Description;

                var hasQuery = !string.IsNullOrWhiteSpace(prm.Query);
                var isDate = string.Equals(prm.Type, "DATE", StringComparison.OrdinalIgnoreCase);
                var isBoolean = IsBooleanType(prm.Type);
                var isNumeric = IsNumericType(prm.Type);

                var valueItemType = isBoolean ? BoFormItemTypes.it_CHECK_BOX : BoFormItemTypes.it_EDIT;
                var valueItem = form.Items.Add(valUid, valueItemType);
                valueItem.Left = baseLeft + 125;
                valueItem.Top = nextTop;
                valueItem.Width = isBoolean ? 20 : 100;

                var context = new ParameterUiContext
                {
                    ValueItemUid = valUid,
                    DescriptionItemUid = descUid,
                    DescriptionQuery = prm.DescriptionQuery,
                    ControlType = isBoolean ? ParameterControlType.CheckBox : ParameterControlType.EditText
                };

                if (isBoolean)
                {
                    var checkBox = (CheckBox)valueItem.Specific;
                    var checkDataSourceUid = ParametersPrefix + "ck_" + suffix;
                    form.DataSources.UserDataSources.Add(checkDataSourceUid, BoDataType.dt_SHORT_TEXT, 1);
                    checkBox.DataBind.SetBound(true, string.Empty, checkDataSourceUid);
                    checkBox.Caption = string.Empty;
                    checkBox.ValOn = "Y";
                    checkBox.ValOff = "N";
                    checkBox.Checked = false;
                }
                else if (isNumeric)
                {
                    ((EditText)valueItem.Specific).Value = string.Empty;
                }

                if (hasQuery)
                {
                    var buttonItem = form.Items.Add(btnUid, BoFormItemTypes.it_BUTTON);
                    buttonItem.Left = baseLeft + 230;
                    buttonItem.Top = nextTop;
                    buttonItem.Width = 24;
                    buttonItem.Height = valueItem.Height;
                    ((Button)buttonItem.Specific).Caption = "...";

                    context.Query = prm.Query;
                    context.ButtonItemUid = btnUid;
                }

                if (prm.ShowDescription && !string.IsNullOrWhiteSpace(prm.DescriptionQuery))
                {
                    var descItem = form.Items.Add(descUid, BoFormItemTypes.it_EDIT);
                    descItem.Left = hasQuery ? baseLeft + 260 : baseLeft + 230;
                    descItem.Top = nextTop;
                    descItem.Width = 250;
                    descItem.Enabled = false;
                }

                if (isDate && !isBoolean)
                {
                    var dateDataSourceUid = ParametersPrefix + "dt_" + suffix;
                    form.DataSources.UserDataSources.Add(dateDataSourceUid, BoDataType.dt_DATE);
                    valueItem.AffectsFormMode = false;
                    ((EditText)valueItem.Specific).DataBind.SetBound(true, string.Empty, dateDataSourceUid);
                    form.DataSources.UserDataSources.Item(dateDataSourceUid).ValueEx = DateTime.Today.ToString("yyyyMMdd");
                }

                _parameterContexts[valUid] = context;
                if (!string.IsNullOrWhiteSpace(context.ButtonItemUid))
                {
                    _parameterContexts[context.ButtonItemUid] = context;
                }

                nextTop += 24;
            }
        }

        private void ResizeMappingFormHeight(Form form, int parameterCount)
        {
            var rowsHeight = Math.Max(parameterCount, 1) * ParameterRowHeight;
            var calculatedHeight = MappingFormBottomPadding + rowsHeight;
            form.Height = Math.Max(MappingFormMinHeight, calculatedHeight);
        }

        private void AddGenerateReportButton(Form form, int parameterCount)
        {
            var rowsHeight = Math.Max(parameterCount, 1) * ParameterRowHeight;
            var buttonTop = 42 + rowsHeight;

            var buttonItem = form.Items.Add(GenerateReportButtonUid, BoFormItemTypes.it_BUTTON);
            buttonItem.Left = 400;
            buttonItem.Top = buttonTop;
            buttonItem.Width = 120;
            ((Button)buttonItem.Specific).Caption = GenerateReportButtonCaption;
        }

        private ReportInfo GetCurrentReportInfo(Form mappingForm)
        {
            if (mappingForm == null || !HasItem(mappingForm, MappingHeaderUid))
            {
                return null;
            }

            var header = ((StaticText)mappingForm.Items.Item(MappingHeaderUid).Specific).Caption;
            if (string.IsNullOrWhiteSpace(header))
            {
                return null;
            }

            var normalized = header.Replace("Reporte:", string.Empty).Trim();
            var separator = " - ";
            var separatorIndex = normalized.IndexOf(separator, StringComparison.OrdinalIgnoreCase);
            if (separatorIndex <= 0)
            {
                return new ReportInfo
                {
                    ReportCode = normalized,
                    ReportName = string.Empty,
                };
            }

            return new ReportInfo
            {
                ReportCode = normalized.Substring(0, separatorIndex).Trim(),
                ReportName = normalized.Substring(separatorIndex + separator.Length).Trim(),
            };
        }

        private string ResolveReportFilePath(string reportCode, string reportName)
        {
            if (string.IsNullOrWhiteSpace(reportCode) || !Directory.Exists(ReportsBasePath))
            {
                return null;
            }

            var expectedFileName = string.IsNullOrWhiteSpace(reportName)
                ? string.Empty
                : $"{reportCode} - {reportName}.rpt";

            if (!string.IsNullOrWhiteSpace(expectedFileName))
            {
                var exactPath = Directory
                    .EnumerateFiles(ReportsBasePath, expectedFileName, SearchOption.AllDirectories)
                    .FirstOrDefault();

                if (!string.IsNullOrWhiteSpace(exactPath))
                {
                    return exactPath;
                }
            }

            return Directory
                .EnumerateFiles(ReportsBasePath, $"{reportCode} - *.rpt", SearchOption.AllDirectories)
                .FirstOrDefault();
        }

        private sealed class ReportInfo
        {
            public string ReportCode { get; set; }
            public string ReportName { get; set; }
        }

        private string ExecuteScalar(string query)
        {
            Recordset recordset = null;
            try
            {
                recordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery(query);
                if (recordset.EoF || recordset.Fields.Count == 0)
                {
                    return string.Empty;
                }

                return Convert.ToString(recordset.Fields.Item(0).Value);
            }
            finally
            {
                if (recordset != null)
                {
                    Marshal.ReleaseComObject(recordset);
                }
            }
        }

        private static string ReplaceFiltro(string query, string selectedValue)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return string.Empty;
            }

            return query.Replace("filtro", (selectedValue ?? string.Empty).Replace("'", "''"));
        }

        private sealed class ReportParameterDefinition
        {
            public string ParamId { get; set; }
            public string Description { get; set; }
            public string Type { get; set; }
            public string Query { get; set; }
            public bool ShowDescription { get; set; }
            public string DescriptionQuery { get; set; }
        }

        private sealed class QueryPickerContext
        {
            public string BaseQuery { get; set; }
            public string SelectedColumn { get; set; }
        }

        private sealed class ParameterUiContext
        {
            public string ValueItemUid { get; set; }
            public string ButtonItemUid { get; set; }
            public string DescriptionItemUid { get; set; }
            public string Query { get; set; }
            public string DescriptionQuery { get; set; }
            public ParameterControlType ControlType { get; set; } = ParameterControlType.EditText;
        }

        private enum ParameterControlType
        {
            EditText,
            CheckBox
        }

    }

}
