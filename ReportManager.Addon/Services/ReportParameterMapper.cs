using ReportManager.Addon.Logging;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Services
{
    public sealed class ReportParameterMapper
    {
        private const int ParametersPane = 1;
        private const string ParametersContainerUid = "Item_8";
        private const string ParametersPrefix = "prm_";
        private const string QueryResultFormType = "RM_QRY_PICK";
        private const string QueryResultGridUid = "grd_qry";
        private const string QueryResultDataTableUid = "DT_QRY";
        private const string QueryResultSourceDataSource = "UD_SRC";

        private readonly Application _app;
        private readonly Logger _log;
        private readonly SAPbobsCOM.Company _company;
        private readonly Dictionary<string, ParameterUiContext> _parameterContexts = new Dictionary<string, ParameterUiContext>(StringComparer.OrdinalIgnoreCase);

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

        public void LoadFromSelectedReportRow(Form form, string reportsGridUid, int selectedRow)
        {
            if (form == null) return;

            try
            {
                var grid = (Grid)form.Items.Item(reportsGridUid).Specific;
                var reportCode = Convert.ToString(grid.DataTable.GetValue("U_SS_IDRPT", selectedRow));
                if (string.IsNullOrWhiteSpace(reportCode))
                {
                    return;
                }

                var parameters = GetReportParameters(reportCode);
                RenderParameters(form, parameters);
            }
            catch (Exception ex)
            {
                _log.Error("No se pudieron mapear los parámetros del reporte.", ex);
                _app.StatusBar.SetText("No se pudieron cargar los parámetros del reporte.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
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
            form.Title = "Seleccionar valor";
            form.Width = 550;
            form.Height = 400;

            var gridItem = form.Items.Add(QueryResultGridUid, BoFormItemTypes.it_GRID);
            gridItem.Left = 10;
            gridItem.Top = 10;
            gridItem.Width = 520;
            gridItem.Height = 330;

            var grid = (Grid)gridItem.Specific;
            var dt = form.DataSources.DataTables.Add(QueryResultDataTableUid);
            dt.ExecuteQuery(context.Query);
            grid.DataTable = dt;
            grid.SelectionMode = BoMatrixSelect.ms_Single;
            grid.AutoResizeColumns();

            var source = sourceFormUid + "|" + context.ValueItemUid;
            form.DataSources.UserDataSources.Add(QueryResultSourceDataSource, BoDataType.dt_LONG_TEXT, 200);
            form.DataSources.UserDataSources.Item(QueryResultSourceDataSource).ValueEx = source;
            form.Visible = true;
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
                ((EditText)sourceForm.Items.Item(valueUid).Specific).Value = selectedValue;

                if (_parameterContexts.TryGetValue(valueUid, out var context)
                    && !string.IsNullOrWhiteSpace(context.DescriptionItemUid)
                    && !string.IsNullOrWhiteSpace(context.DescriptionQuery))
                {
                    var description = ExecuteScalar(ReplaceFiltro(context.DescriptionQuery, selectedValue));
                    ((EditText)sourceForm.Items.Item(context.DescriptionItemUid).Specific).Value = description;
                }

                queryForm.Close();
            }
            catch (Exception ex)
            {
                _log.Error("No se pudo aplicar la selección de consulta.", ex);
            }
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
            ClearParameterControls(form);
            if (parameters.Count == 0)
            {
                return;
            }

            var container = form.Items.Item(ParametersContainerUid);
            var baseLeft = container.Left + 10;
            var baseTop = container.Top + 10;
            var nextTop = baseTop;

            for (int i = 0; i < parameters.Count; i++)
            {
                var prm = parameters[i];
                var suffix = i.ToString("00");
                var lblUid = ParametersPrefix + "lbl_" + suffix;
                var valUid = ParametersPrefix + "val_" + suffix;
                var btnUid = ParametersPrefix + "btn_" + suffix;
                var descUid = ParametersPrefix + "dsc_" + suffix;

                var labelItem = form.Items.Add(lblUid, BoFormItemTypes.it_STATIC);
                labelItem.FromPane = ParametersPane;
                labelItem.ToPane = ParametersPane;
                labelItem.Left = baseLeft;
                labelItem.Top = nextTop + 2;
                labelItem.Width = 120;
                ((StaticText)labelItem.Specific).Caption = string.IsNullOrWhiteSpace(prm.Description) ? prm.ParamId : prm.Description;

                var valueItem = form.Items.Add(valUid, BoFormItemTypes.it_EDIT);
                valueItem.FromPane = ParametersPane;
                valueItem.ToPane = ParametersPane;
                valueItem.Left = baseLeft + 125;
                valueItem.Top = nextTop;
                valueItem.Width = 120;

                var context = new ParameterUiContext
                {
                    ValueItemUid = valUid,
                    DescriptionItemUid = descUid,
                    DescriptionQuery = prm.DescriptionQuery
                };

                var hasQuery = !string.IsNullOrWhiteSpace(prm.Query);
                var isDate = string.Equals(prm.Type, "DATE", StringComparison.OrdinalIgnoreCase);

                if (hasQuery)
                {
                    var buttonItem = form.Items.Add(btnUid, BoFormItemTypes.it_BUTTON);
                    buttonItem.FromPane = ParametersPane;
                    buttonItem.ToPane = ParametersPane;
                    buttonItem.Left = baseLeft + 248;
                    buttonItem.Top = nextTop;
                    buttonItem.Width = 22;
                    ((Button)buttonItem.Specific).Caption = "...";

                    context.Query = prm.Query;
                    context.ButtonItemUid = btnUid;
                }

                if (prm.ShowDescription && !string.IsNullOrWhiteSpace(prm.DescriptionQuery))
                {
                    var descItem = form.Items.Add(descUid, BoFormItemTypes.it_EDIT);
                    descItem.FromPane = ParametersPane;
                    descItem.ToPane = ParametersPane;
                    descItem.Left = hasQuery ? baseLeft + 275 : baseLeft + 248;
                    descItem.Top = nextTop;
                    descItem.Width = 140;
                    descItem.Enabled = false;
                }

                if (isDate)
                {
                    ((EditText)valueItem.Specific).Value = DateTime.Today.ToString("yyyyMMdd");
                }

                _parameterContexts[valUid] = context;
                if (!string.IsNullOrWhiteSpace(context.ButtonItemUid))
                {
                    _parameterContexts[context.ButtonItemUid] = context;
                }

                nextTop += 24;
            }
        }

        private void ClearParameterControls(Form form)
        {
            var dynamicItemUids = _parameterContexts.Values
                .SelectMany(x => new[] { x.ValueItemUid, x.ButtonItemUid, x.DescriptionItemUid })
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var uid in dynamicItemUids)
            {
                RemoveItemIfExists(form, uid);
            }

            for (int i = 0; i < 100; i++)
            {
                RemoveItemIfExists(form, ParametersPrefix + "lbl_" + i.ToString("00"));
            }

            _parameterContexts.Clear();
        }

        private static void RemoveItemIfExists(Form form, string itemUid)
        {
            try
            {
                form.Items.Item(itemUid);
                //form.Items.Remove(itemUid);
            }
            catch
            {
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

        private sealed class ReportParameterDefinition
        {
            public string ParamId { get; set; }
            public string Description { get; set; }
            public string Type { get; set; }
            public string Query { get; set; }
            public bool ShowDescription { get; set; }
            public string DescriptionQuery { get; set; }
        }

        private sealed class ParameterUiContext
        {
            public string ValueItemUid { get; set; }
            public string ButtonItemUid { get; set; }
            public string DescriptionItemUid { get; set; }
            public string Query { get; set; }
            public string DescriptionQuery { get; set; }
        }
    }

}
