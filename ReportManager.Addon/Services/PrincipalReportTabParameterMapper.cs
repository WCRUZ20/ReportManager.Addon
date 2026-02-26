using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using ReportManager.Addon.Entidades;
using ReportManager.Addon.Forms;
using ReportManager.Addon.Logging;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace ReportManager.Addon.Services
{
    public sealed class PrincipalReportTabParameterMapper
    {
        private const string TabControlUid = "tbc_prm";
        private const string ReportsGridUid = "grd_rpt";
        private const string ParametersPrefix = "tp";
        private const string GenerateButtonPrefix = "tg";
        private const string GenerateButtonCaption = "Generar reporte";
        private const string QueryResultFormType = "RM_QRY_TBC";
        private const string QueryResultGridUid = "grd_qry";
        private const string QueryResultDataTableUid = "DT_QRY";
        private const string ReportsBasePath = @"C:\Reportes SAP";

        private readonly Application _app;
        private readonly Logger _log;
        private readonly SAPbobsCOM.Company _company;
        private readonly Dictionary<string, ReportTabContext> _tabsByReportCode = new Dictionary<string, ReportTabContext>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ReportTabContext> _tabsByGenerateButton = new Dictionary<string, ReportTabContext>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ParameterUiContext> _parameterContextsByButton = new Dictionary<string, ParameterUiContext>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, QueryPickerContext> _queryPickerContexts = new Dictionary<string, QueryPickerContext>(StringComparer.OrdinalIgnoreCase);
        private int _nextPaneLevel = 2;

        public PrincipalReportTabParameterMapper(Application app, Logger log, SAPbobsCOM.Company company)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _company = company ?? throw new ArgumentNullException(nameof(company));
        }

        public bool IsParameterButton(string itemUid)
        {
            return !string.IsNullOrWhiteSpace(itemUid)
                && itemUid.StartsWith(ParametersPrefix + "b", StringComparison.OrdinalIgnoreCase);
        }

        public bool IsGenerateReportButton(string itemUid)
        {
            return !string.IsNullOrWhiteSpace(itemUid)
                && itemUid.StartsWith(GenerateButtonPrefix, StringComparison.OrdinalIgnoreCase);
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

        public void ShowFromSelectedReportRow(Form principalForm, int selectedRow)
        {
            if (principalForm == null)
            {
                return;
            }

            try
            {
                var grid = (Grid)principalForm.Items.Item(ReportsGridUid).Specific;
                var reportCode = Convert.ToString(grid.DataTable.GetValue("U_SS_IDRPT", selectedRow));
                var reportName = Convert.ToString(grid.DataTable.GetValue("U_SS_NOMBRPT", selectedRow));
                if (string.IsNullOrWhiteSpace(reportCode))
                {
                    return;
                }

                ReportTabContext tabContext;
                if (!_tabsByReportCode.TryGetValue(reportCode, out tabContext))
                {
                    var parameters = GetReportParameters(reportCode);
                    tabContext = CreateReportTab(principalForm, reportCode, reportName, parameters);
                    _tabsByReportCode[reportCode] = tabContext;
                    _tabsByGenerateButton[tabContext.GenerateButtonUid] = tabContext;
                }

                ActivateTab(principalForm, tabContext);
            }
            catch (Exception ex)
            {
                _log.Error("No se pudieron mapear los parámetros del reporte en tab principal.", ex);
                _app.StatusBar.SetText("No se pudieron cargar los parámetros del reporte.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public void GenerateSelectedReport(string principalFormUid, string generateButtonUid)
        {
            if (string.IsNullOrWhiteSpace(principalFormUid) || string.IsNullOrWhiteSpace(generateButtonUid))
            {
                return;
            }

            ReportTabContext tabContext;
            if (!_tabsByGenerateButton.TryGetValue(generateButtonUid, out tabContext))
            {
                return;
            }

            try
            {
                var form = _app.Forms.Item(principalFormUid);
                if (form == null)
                {
                    return;
                }

                if (!TryValidateRequiredParameters(form, tabContext, out var validationMessage))
                {
                    _app.StatusBar.SetText(validationMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }

                var reportFilePath = ResolveReportFilePath(tabContext.ReportCode, tabContext.ReportName);
                if (string.IsNullOrWhiteSpace(reportFilePath))
                {
                    _app.StatusBar.SetText("No se encontró el archivo de Crystal Report para el reporte seleccionado.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }

                var parameterValues = BuildParameterValues(form, tabContext);
                OpenCrystalViewerOnStaThread(reportFilePath, parameterValues);
                _app.StatusBar.SetText("Reporte abierto correctamente.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo abrir el Crystal Report seleccionado desde tab principal.", ex);
                _app.StatusBar.SetText("No se pudo abrir el reporte seleccionado.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public void OpenQuerySelector(string sourceFormUid, string buttonItemUid)
        {
            ParameterUiContext context;
            if (!_parameterContextsByButton.TryGetValue(buttonItemUid, out context) || string.IsNullOrWhiteSpace(context.Query))
            {
                return;
            }

            var queryFormUid = QueryResultFormType + DateTime.Now.Ticks;
            var creationParams = (FormCreationParams)_app.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            creationParams.UniqueID = queryFormUid;
            creationParams.FormType = QueryResultFormType;
            creationParams.BorderStyle = BoFormBorderStyle.fbs_Sizable;

            var form = _app.Forms.AddEx(creationParams);
            form.Title = "Seleccionar valor";
            form.Width = 550;
            form.Height = 400;

            var gridItem = form.Items.Add(QueryResultGridUid, BoFormItemTypes.it_GRID);
            gridItem.Left = 10;
            gridItem.Top = 10;
            gridItem.Width = 500;
            gridItem.Height = 330;

            var grid = (Grid)gridItem.Specific;
            var dt = form.DataSources.DataTables.Add(QueryResultDataTableUid);
            dt.ExecuteQuery(context.Query);
            grid.DataTable = dt;
            grid.SelectionMode = BoMatrixSelect.ms_Single;
            grid.Item.Enabled = false;
            grid.AutoResizeColumns();

            _queryPickerContexts[queryFormUid] = new QueryPickerContext
            {
                SourceFormUid = sourceFormUid,
                ValueItemUid = context.ValueItemUid,
                DescriptionItemUid = context.DescriptionItemUid,
                DescriptionQuery = context.DescriptionQuery
            };

            form.Visible = true;
        }

        public void ApplyQuerySelection(string queryFormUid, int selectedRow)
        {
            QueryPickerContext context;
            if (!_queryPickerContexts.TryGetValue(queryFormUid, out context))
            {
                return;
            }

            try
            {
                var queryForm = _app.Forms.Item(queryFormUid);
                var grid = (Grid)queryForm.Items.Item(QueryResultGridUid).Specific;
                if (grid.DataTable.Columns.Count == 0)
                {
                    return;
                }

                var selectedColumnName = grid.DataTable.Columns.Item(0).Name;
                var value = Convert.ToString(grid.DataTable.GetValue(selectedColumnName, selectedRow));

                var sourceForm = _app.Forms.Item(context.SourceFormUid);
                ((EditText)sourceForm.Items.Item(context.ValueItemUid).Specific).Value = value;

                if (!string.IsNullOrWhiteSpace(context.DescriptionItemUid)
                    && HasItem(sourceForm, context.DescriptionItemUid)
                    && !string.IsNullOrWhiteSpace(context.DescriptionQuery))
                {
                    var resolvedDescription = ExecuteScalar(ReplaceFiltro(context.DescriptionQuery, value));
                    ((EditText)sourceForm.Items.Item(context.DescriptionItemUid).Specific).Value = resolvedDescription;
                }

                _queryPickerContexts.Remove(queryFormUid);
                queryForm.Close();
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo aplicar la selección de consulta del tab principal.", ex);
            }
        }

        private ReportTabContext CreateReportTab(Form principalForm, string reportCode, string reportName, List<ReportParameterDefinition> parameters)
        {
            var paneLevel = _nextPaneLevel++;
            var folderUid = BuildFolderUid(_tabsByReportCode.Count + 1);

            Item folderItem;
            if (_tabsByReportCode.Count == 0)
            {
                folderUid = TabControlUid;
                folderItem = principalForm.Items.Item(TabControlUid);
            }
            else
            {
                folderItem = principalForm.Items.Add(folderUid, BoFormItemTypes.it_FOLDER);
                var anchor = principalForm.Items.Item(TabControlUid);
                folderItem.Top = anchor.Top;
                folderItem.Left = anchor.Left + (_tabsByReportCode.Count * 102);
                folderItem.Width = 100;
                folderItem.Height = 20;
                var folder = (Folder)folderItem.Specific;
                folder.GroupWith(TabControlUid);
            }

            var folderSpecific = (Folder)folderItem.Specific;
            folderSpecific.Caption = reportCode;
            folderSpecific.Pane = paneLevel;

            var tabContext = new ReportTabContext
            {
                ReportCode = reportCode,
                ReportName = reportName,
                FolderUid = folderUid,
                PaneLevel = paneLevel,
                ParameterInstances = new List<ParameterInstance>()
            };

            var top = 70;
            for (int i = 0; i < parameters.Count; i++)
            {
                var prm = parameters[i];
                var suffix = paneLevel.ToString() + i.ToString("00");
                var labelUid = ParametersPrefix + "l" + suffix;
                var valueUid = ParametersPrefix + "v" + suffix;
                var buttonUid = ParametersPrefix + "b" + suffix;
                var descUid = ParametersPrefix + "d" + suffix;

                var label = principalForm.Items.Add(labelUid, BoFormItemTypes.it_STATIC);
                label.FromPane = paneLevel;
                label.ToPane = paneLevel;
                label.Left = 430;
                label.Top = top + 2;
                label.Width = 120;
                ((StaticText)label.Specific).Caption = string.IsNullOrWhiteSpace(prm.Description) ? prm.ParamId : prm.Description;

                var hasQuery = !string.IsNullOrWhiteSpace(prm.Query);
                var isBoolean = IsBooleanType(prm.Type);
                var isNumeric = IsNumericType(prm.Type);
                var isDate = string.Equals(prm.Type, "DATE", StringComparison.OrdinalIgnoreCase);

                var valueItem = principalForm.Items.Add(valueUid, isBoolean ? BoFormItemTypes.it_CHECK_BOX : BoFormItemTypes.it_EDIT);
                valueItem.FromPane = paneLevel;
                valueItem.ToPane = paneLevel;
                valueItem.Left = 560;
                valueItem.Top = top;
                valueItem.Width = isBoolean ? 20 : 100;

                var context = new ParameterUiContext
                {
                    ValueItemUid = valueUid,
                    DescriptionItemUid = descUid,
                    DescriptionQuery = prm.DescriptionQuery,
                    Query = prm.Query,
                    ParameterType = prm.Type
                };

                if (isBoolean)
                {
                    var checkBox = (CheckBox)valueItem.Specific;
                    var checkDataSourceUid = ParametersPrefix + "c" + suffix;
                    EnsureUserDataSource(principalForm, checkDataSourceUid, BoDataType.dt_SHORT_TEXT, 1);
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

                if (isDate && !isBoolean)
                {
                    var dateDataSourceUid = ParametersPrefix + "t" + suffix;
                    EnsureUserDataSource(principalForm, dateDataSourceUid, BoDataType.dt_DATE);
                    ((EditText)valueItem.Specific).DataBind.SetBound(true, string.Empty, dateDataSourceUid);
                    principalForm.DataSources.UserDataSources.Item(dateDataSourceUid).ValueEx = DateTime.Today.ToString("yyyyMMdd");
                }

                if (hasQuery)
                {
                    var queryButton = principalForm.Items.Add(buttonUid, BoFormItemTypes.it_BUTTON);
                    queryButton.FromPane = paneLevel;
                    queryButton.ToPane = paneLevel;
                    queryButton.Left = 665;
                    queryButton.Top = top;
                    queryButton.Width = 24;
                    queryButton.Height = valueItem.Height;
                    ((Button)queryButton.Specific).Caption = "...";

                    _parameterContextsByButton[buttonUid] = context;
                }

                if (prm.ShowDescription && !string.IsNullOrWhiteSpace(prm.DescriptionQuery))
                {
                    var desc = principalForm.Items.Add(descUid, BoFormItemTypes.it_EDIT);
                    desc.FromPane = paneLevel;
                    desc.ToPane = paneLevel;
                    desc.Left = 694;
                    desc.Top = top;
                    desc.Width = 250;
                    desc.Enabled = false;
                }

                tabContext.ParameterInstances.Add(new ParameterInstance
                {
                    Definition = prm,
                    ValueItemUid = valueUid
                });

                top += 24;
            }

            var generateUid = GenerateButtonPrefix + paneLevel.ToString("00000000");
            var generateButton = principalForm.Items.Add(generateUid, BoFormItemTypes.it_BUTTON);
            generateButton.FromPane = paneLevel;
            generateButton.ToPane = paneLevel;
            generateButton.Left = 824;
            generateButton.Top = Math.Max(70, top + 10);
            generateButton.Width = 120;
            ((Button)generateButton.Specific).Caption = GenerateButtonCaption;
            tabContext.GenerateButtonUid = generateUid;

            return tabContext;
        }

        private void ActivateTab(Form principalForm, ReportTabContext selected)
        {
            foreach (var entry in _tabsByReportCode.Values)
            {
                if (!HasItem(principalForm, entry.FolderUid))
                {
                    continue;
                }

                principalForm.Items.Item(entry.FolderUid).Visible = string.Equals(entry.ReportCode, selected.ReportCode, StringComparison.OrdinalIgnoreCase);
            }

            principalForm.PaneLevel = selected.PaneLevel;
            if (HasItem(principalForm, selected.FolderUid))
            {
                ((Folder)principalForm.Items.Item(selected.FolderUid).Specific).Select();
            }
        }

        private Dictionary<string, object> BuildParameterValues(Form principalForm, ReportTabContext tabContext)
        {
            var values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < tabContext.ParameterInstances.Count; i++)
            {
                var instance = tabContext.ParameterInstances[i];
                if (instance?.Definition == null || string.IsNullOrWhiteSpace(instance.Definition.ParamId) || !HasItem(principalForm, instance.ValueItemUid))
                {
                    continue;
                }

                var value = GetUiValue(principalForm, instance.ValueItemUid, instance.Definition.Type);
                if (value != null)
                {
                    values[instance.Definition.ParamId] = value;
                }
            }

            return values;
        }

        private bool TryValidateRequiredParameters(Form principalForm, ReportTabContext tabContext, out string validationMessage)
        {
            validationMessage = string.Empty;
            for (int i = 0; i < tabContext.ParameterInstances.Count; i++)
            {
                var instance = tabContext.ParameterInstances[i];
                if (instance?.Definition == null || !instance.Definition.IsRequired || !HasItem(principalForm, instance.ValueItemUid))
                {
                    continue;
                }

                var value = GetUiValue(principalForm, instance.ValueItemUid, instance.Definition.Type);
                var isMissing = value == null || (value is string stringValue && string.IsNullOrWhiteSpace(stringValue));
                if (isMissing)
                {
                    var parameterName = !string.IsNullOrWhiteSpace(instance.Definition.Description)
                        ? instance.Definition.Description
                        : instance.Definition.ParamId;
                    validationMessage = $"El parámetro obligatorio '{parameterName}' debe ser diligenciado.";
                    return false;
                }
            }

            return true;
        }

        private object GetUiValue(Form principalForm, string valueUid, string parameterType)
        {
            if (IsBooleanType(parameterType))
            {
                var checkBox = (CheckBox)principalForm.Items.Item(valueUid).Specific;
                return checkBox.Checked;
            }

            var rawValue = ((EditText)principalForm.Items.Item(valueUid).Specific).Value;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return null;
            }

            if (string.Equals(parameterType, "DATE", StringComparison.OrdinalIgnoreCase)
                && DateTime.TryParse(rawValue, out var dateValue))
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

        private void OpenCrystalViewerOnStaThread(string reportFilePath, Dictionary<string, object> parameterValues)
        {
            var viewerThread = new Thread(() =>
            {
                ReportDocument localReportDocument = null;
                try
                {
                    localReportDocument = new ReportDocument();
                    localReportDocument.Load(reportFilePath);
                    localReportDocument.SetDatabaseLogon(Globals.dbuser, Globals.pwduser);
                    ApplyParameters(localReportDocument, parameterValues);

                    System.Windows.Forms.Application.Run(new CrystalReportViewerForm(localReportDocument));
                }
                catch (Exception ex)
                {
                    _log.Error("No se pudo abrir el visor de Crystal Reports en hilo STA (tabs principal).", ex);
                    localReportDocument?.Dispose();
                }
            });

            viewerThread.IsBackground = true;
            viewerThread.SetApartmentState(ApartmentState.STA);
            viewerThread.Start();
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

        private List<ReportParameterDefinition> GetReportParameters(string reportCode)
        {
            var result = new List<ReportParameterDefinition>();
            Recordset recordset = null;

            try
            {
                var escapedReportCode = reportCode.Replace("'", "''");
                var query = _company.DbServerType == BoDataServerTypes.dst_HANADB
                    ? $@"select T1.""LineId"", T1.""U_SS_IDPARAM"", T1.""U_SS_DSCPARAM"", T1.""U_SS_TIPO"", T1.""U_SS_OBLIGA"", T1.""U_SS_QUERY"", T1.""U_SS_DESC"", T1.""U_SS_QUERYD"", T1.""U_SS_ACTIVO""
                        from ""@SS_PRM_CAB"" T0
                        inner join ""@SS_PRM_DET"" T1 on T0.""Code"" = T1.""Code""
                        where T0.""U_SS_IDRPT"" = '{escapedReportCode}'
                        order by T1.""LineId"""
                    : $@"select T1.LineId, T1.U_SS_IDPARAM, T1.U_SS_DSCPARAM, T1.U_SS_TIPO, T1.U_SS_OBLIGA, T1.U_SS_QUERY, T1.U_SS_DESC, T1.U_SS_QUERYD, T1.U_SS_ACTIVO
                        from [@SS_PRM_CAB] T0
                        inner join [@SS_PRM_DET] T1 on T0.Code = T1.Code
                        where T0.U_SS_IDRPT = '{escapedReportCode}'
                        order by T1.LineId";

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
                            IsRequired = string.Equals(Convert.ToString(recordset.Fields.Item("U_SS_OBLIGA").Value), "Y", StringComparison.OrdinalIgnoreCase),
                            Query = Convert.ToString(recordset.Fields.Item("U_SS_QUERY").Value),
                            ShowDescription = string.Equals(Convert.ToString(recordset.Fields.Item("U_SS_DESC").Value), "Y", StringComparison.OrdinalIgnoreCase),
                            DescriptionQuery = Convert.ToString(recordset.Fields.Item("U_SS_QUERYD").Value)
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

        private static string BuildFolderUid(int ordinal)
        {
            return "tb" + ordinal.ToString("00000000");
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

        private static void EnsureUserDataSource(Form form, string dataSourceUid, BoDataType dataType, int size = 0)
        {
            try
            {
                form.DataSources.UserDataSources.Item(dataSourceUid);
            }
            catch
            {
                if (size > 0)
                {
                    form.DataSources.UserDataSources.Add(dataSourceUid, dataType, size);
                }
                else
                {
                    form.DataSources.UserDataSources.Add(dataSourceUid, dataType);
                }
            }
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

        private sealed class ReportTabContext
        {
            public string ReportCode { get; set; }
            public string ReportName { get; set; }
            public string FolderUid { get; set; }
            public int PaneLevel { get; set; }
            public string GenerateButtonUid { get; set; }
            public List<ParameterInstance> ParameterInstances { get; set; }
        }

        private sealed class ParameterInstance
        {
            public ReportParameterDefinition Definition { get; set; }
            public string ValueItemUid { get; set; }
        }

        private sealed class ParameterUiContext
        {
            public string ValueItemUid { get; set; }
            public string DescriptionItemUid { get; set; }
            public string Query { get; set; }
            public string DescriptionQuery { get; set; }
            public string ParameterType { get; set; }
        }

        private sealed class QueryPickerContext
        {
            public string SourceFormUid { get; set; }
            public string ValueItemUid { get; set; }
            public string DescriptionItemUid { get; set; }
            public string DescriptionQuery { get; set; }
        }

        private sealed class ReportParameterDefinition
        {
            public string ParamId { get; set; }
            public string Description { get; set; }
            public string Type { get; set; }
            public bool IsRequired { get; set; }
            public string Query { get; set; }
            public bool ShowDescription { get; set; }
            public string DescriptionQuery { get; set; }
        }
    }
}
