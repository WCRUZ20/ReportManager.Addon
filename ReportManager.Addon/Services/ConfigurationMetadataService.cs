using ReportManager.Addon.Logging;
using ReportManager.Addon.Models;
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
    public sealed class ConfigurationMetadataService
    {
        private readonly Application _app;
        private readonly Logger _log;
        private readonly SAPbobsCOM.Company _company;

        public ConfigurationMetadataService(Application app, Logger log, SAPbobsCOM.Company company)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _company = company ?? throw new ArgumentNullException(nameof(company));
        }

        public void CreateDepartmentTable()
        {
            CreateUserTableIfNotExists("SS_DPTS", "Departamentos", BoUTBTableType.bott_NoObject);
            _app.StatusBar.SetText("Estructura SS_DPTS validada.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        public void CreateParamterTypeTable()
        {
            CreateUserTableIfNotExists("SS_PRMTYPE", "Tipos parámetros", BoUTBTableType.bott_NoObject);
            _app.StatusBar.SetText("Estructura SS_PRMTYPE validada.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        public void CreateReportConfigurationStructures()
        {
            CreateParameterStructures();
            //CreateDefinitionStructures();
            _app.StatusBar.SetText("Estructuras maestras y UDOs validadas.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        private void CreateParameterStructures()
        {
            CreateUserTableIfNotExists("SS_PRM_CAB", "Parametros Reporte Cab", BoUTBTableType.bott_MasterData);
            CreateUserFieldIfNotExists("@SS_PRM_CAB", "SS_IDRPT", "Id Reporte", BoFieldTypes.db_Alpha, 50);
            CreateUserFieldIfNotExists("@SS_PRM_CAB", "SS_NOMBRPT", "Nombre Reporte", BoFieldTypes.db_Alpha, 50);
            CreateUserFieldIfNotExists("@SS_PRM_CAB", "SS_IDDPT", "Id Departamento", BoFieldTypes.db_Alpha, 50, BoFldSubTypes.st_None, null, null, "SS_DPTS");
            CreateUserFieldIfNotExists("@SS_PRM_CAB", "SS_ACTIVO", "Rpt Activo", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "Y", "N");

            CreateUserTableIfNotExists("SS_PRM_DET", "Parametros Reporte Det", BoUTBTableType.bott_MasterDataLines);
            CreateUserFieldIfNotExists("@SS_PRM_DET", "SS_IDPARAM", "Id Parametro", BoFieldTypes.db_Alpha, 50);
            CreateUserFieldIfNotExists("@SS_PRM_DET", "SS_DSCPARAM", "Desc Parametro", BoFieldTypes.db_Alpha, 50);
            CreateUserFieldIfNotExists("@SS_PRM_DET", "SS_TIPO", "Tipo parámetro", BoFieldTypes.db_Alpha, 50, BoFldSubTypes.st_None, null, null, "SS_PRMTYPE");
            CreateUserFieldIfNotExists("@SS_PRM_DET", "SS_OBLIGA", "Obligatorio", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "Y", "N");
            CreateUserFieldIfNotExists("@SS_PRM_DET", "SS_QUERY", "Query", BoFieldTypes.db_Memo, 250);
            CreateUserFieldIfNotExists("@SS_PRM_DET", "SS_DESC", "Descripción selección", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "Y", "N");
            CreateUserFieldIfNotExists("@SS_PRM_DET", "SS_QUERYD", "Query descripción", BoFieldTypes.db_Memo, 250);
            CreateUserFieldIfNotExists("@SS_PRM_DET", "SS_ACTIVO", "Activo", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "Y", "N");

            var def = new UdoDefinition
            {
                Code = "SS_PRM_CAB",
                Name = "Parametrización de reportes",
                TableName = "SS_PRM_CAB",
                ObjectType = BoUDOObjType.boud_MasterData,

                CanFind = BoYesNoEnum.tYES,
                CanCancel = BoYesNoEnum.tNO,
                CanClose = BoYesNoEnum.tNO,
                CanDelete = BoYesNoEnum.tYES,
                CanCreateDefaultForm = BoYesNoEnum.tYES,
                ManageSeries = BoYesNoEnum.tNO,
                UseEnhancedFormColumns = true,
                RebuildEnhancedForm = BoYesNoEnum.tYES,
                EnableEnhancedForm = BoYesNoEnum.tYES,
            };

            def.FindColumns.Add(new UdoFindColumn { Alias = "Code", Description = "Código" });
            def.FindColumns.Add(new UdoFindColumn { Alias = "U_SS_IDRPT", Description = "Id Reporte" });
            def.FindColumns.Add(new UdoFindColumn { Alias = "U_SS_NOMBRPT", Description = "Nombre Reporte" });

            def.ChildTables.Add(new UdoChildTable { TableName = "SS_PRM_DET" });

            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "Code", Description = "Code", ChildNumber = 0, ColumnNumber = 1 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "Name", Description = "Name", ChildNumber = 0, ColumnNumber = 2 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_IDRPT", Description = "Id Reporte", ChildNumber = 0, ColumnNumber = 3 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_NOMBRPT", Description = "Nombre Reporte", ChildNumber = 0, ColumnNumber = 4 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_IDDPT", Description = "Departamento", ChildNumber = 0, ColumnNumber = 5 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_ACTIVO", Description = "Rpt Activo", ChildNumber = 0, ColumnNumber = 6 });

            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_IDPARAM", Description = "Id Param", ChildNumber = 1, ColumnNumber = 1 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_DSCPARAM", Description = "Descripción", ChildNumber = 1, ColumnNumber = 2 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_TIPO", Description = "Tipo parámetro", ChildNumber = 1, ColumnNumber = 3 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_OBLIGA", Description = "Obligatorio", ChildNumber = 1, ColumnNumber = 4 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_QUERY", Description = "Query", ChildNumber = 1, ColumnNumber = 5 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_DESC", Description = "Descripción selección", ChildNumber = 1, ColumnNumber = 6 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_QUERYD", Description = "Query Descripción", ChildNumber = 1, ColumnNumber = 7 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_ACTIVO", Description = "Activo", ChildNumber = 1, ColumnNumber = 8 });

            RegisterUdoIfNotExists(def);
        }

        private void CreateDefinitionStructures()
        {
            CreateUserTableIfNotExists("SS_DFRPT_CAB", "Definicion Reporte Cab", BoUTBTableType.bott_MasterData);
            CreateUserFieldIfNotExists("@SS_DFRPT_CAB", "SS_IDDPT", "Id Departamento", BoFieldTypes.db_Alpha, 50, BoFldSubTypes.st_None, null, null, "SS_DPTS");

            CreateUserTableIfNotExists("SS_DFRPT_DET", "Definicion Reporte Det", BoUTBTableType.bott_MasterDataLines);
            CreateUserFieldIfNotExists("@SS_DFRPT_DET", "SS_IDRPT", "Id Reporte", BoFieldTypes.db_Alpha, 50, BoFldSubTypes.st_None, null, null, "SS_PRM_CAB");
            CreateUserFieldIfNotExists("@SS_DFRPT_DET", "SS_ACTIVO", "Activo", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "Y", "N");

            var def = new UdoDefinition
            {
                Code = "SS_DFRPT_CAB",
                Name = "Definición de reportes",
                TableName = "SS_DFRPT_CAB",
                ObjectType = BoUDOObjType.boud_MasterData,

                CanFind = BoYesNoEnum.tYES,
                CanCancel = BoYesNoEnum.tNO,
                CanClose = BoYesNoEnum.tNO,
                CanDelete = BoYesNoEnum.tYES,
                CanCreateDefaultForm = BoYesNoEnum.tYES,
                UseEnhancedFormColumns = true,
                RebuildEnhancedForm = BoYesNoEnum.tYES,
                EnableEnhancedForm = BoYesNoEnum.tYES,
            };

            def.FindColumns.Add(new UdoFindColumn { Alias = "Code", Description = "Código" });
            def.FindColumns.Add(new UdoFindColumn { Alias = "U_SS_IDDPT", Description = "Departamento" });

            def.ChildTables.Add(new UdoChildTable { TableName = "SS_DFRPT_DET" });

            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "Code", Description = "Código", ChildNumber = 0, ColumnNumber = 1 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "Name", Description = "Nombre", ChildNumber = 0, ColumnNumber = 2 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_IDDPT", Description = "Departamento", ChildNumber = 0, ColumnNumber = 3 });

            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_IDRPT", Description = "Reporte", ChildNumber = 1, ColumnNumber = 1 });
            def.EnhancedFormColumns.Add(new UdoEnhancedFormColumn { Alias = "U_SS_ACTIVO", Description = "Activo", ChildNumber = 1, ColumnNumber = 2 });

            RegisterUdoIfNotExists(def);
        }

        private void  CreateUserTableIfNotExists(string tableName, string description, BoUTBTableType type)
        {
            UserTablesMD userTables = null;
            Recordset recordset = null;

            try
            {
                //var company = GetCompany();
                recordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery($"SELECT TOP 1 1 FROM OUTB WHERE TableName = '" + tableName + "'");

                if (!recordset.EoF)
                {
                    return;
                }

                ReleaseComObject(recordset);
                recordset = null;

                userTables = (UserTablesMD)_company.GetBusinessObject(BoObjectTypes.oUserTables);
                userTables.TableName = tableName;
                userTables.TableDescription = description;
                userTables.TableType = type;

                AddMetadata(userTables, "No se pudo crear tabla " + tableName + ".");
                _log.Info("Tabla creada: " + tableName);
            }
            finally
            {
                ReleaseComObject(recordset);
                ReleaseComObject(userTables);
            }
        }

        private void CreateUserFieldIfNotExists(
            string tableName,
            string fieldName,
            string description,
            BoFieldTypes fieldType,
            int size,
            BoFldSubTypes subType = BoFldSubTypes.st_None,
            string validValueYes = null,
            string validValueNo = null,
            string linkedTable = null)
        {
            UserFieldsMD fields = null;
            Recordset recordset = null;

            try
            {
                //var company = GetCompany();
                recordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery("SELECT TOP 1 1 FROM CUFD WHERE TableID = '" + tableName + "' AND AliasID = '" + fieldName + "'");

                if (!recordset.EoF)
                {
                    return;
                }

                ReleaseComObject(recordset);
                recordset = null;

                fields = (UserFieldsMD)_company.GetBusinessObject(BoObjectTypes.oUserFields);
                fields.TableName = tableName;
                fields.Name = fieldName;
                fields.Description = description;
                fields.Type = fieldType;

                if (fieldType == BoFieldTypes.db_Alpha)
                {
                    fields.EditSize = size;
                }

                if (subType != BoFldSubTypes.st_None)
                {
                    fields.SubType = subType;
                }

                if (!string.IsNullOrWhiteSpace(linkedTable))
                {
                    fields.LinkedTable = linkedTable.TrimStart('@');
                }

                if (!string.IsNullOrWhiteSpace(validValueYes) && !string.IsNullOrWhiteSpace(validValueNo))
                {
                    fields.ValidValues.Value = validValueYes;
                    fields.ValidValues.Description = "Sí";
                    fields.ValidValues.Add();
                    fields.ValidValues.Value = validValueNo;
                    fields.ValidValues.Description = "No";
                    fields.ValidValues.Add();
                    fields.DefaultValue = validValueYes;
                }

                AddMetadata(fields, "No se pudo crear campo " + fieldName + " en " + tableName + ".");
                _log.Info("Campo creado: " + tableName + "." + fieldName);
            }
            finally
            {
                ReleaseComObject(recordset);
                ReleaseComObject(fields);
            }
        }

        private void RegisterUdoIfNotExists(UdoDefinition def)
        {
            if (def == null) throw new ArgumentNullException(nameof(def));
            if (string.IsNullOrWhiteSpace(def.Code)) throw new ArgumentException("Code requerido.");
            if (string.IsNullOrWhiteSpace(def.Name)) throw new ArgumentException("Name requerido.");
            if (string.IsNullOrWhiteSpace(def.TableName)) throw new ArgumentException("TableName requerido.");

            UserObjectsMD udo = null;
            Recordset rs = null;

            try
            {
                rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery("SELECT TOP 1 1 FROM OUDO WHERE Code = '" + def.Code.Replace("'", "''") + "'");
                if (!rs.EoF) return;

                ReleaseComObject(rs);
                rs = null;

                udo = (UserObjectsMD)_company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

                udo.Code = def.Code;
                udo.Name = def.Name;
                udo.ObjectType = def.ObjectType;
                udo.TableName = def.TableName; // SIN @

                udo.CanFind = def.CanFind;
                udo.CanCancel = def.CanCancel;
                udo.CanClose = def.CanClose;
                udo.CanDelete = def.CanDelete;
                udo.CanCreateDefaultForm = def.CanCreateDefaultForm;
                udo.CanYearTransfer = def.CanYearTransfer;
                udo.ManageSeries = def.ManageSeries;

                if (def.UseLog == BoYesNoEnum.tYES)
                {
                    udo.CanLog = BoYesNoEnum.tYES;
                    if (!string.IsNullOrWhiteSpace(def.LogTableName))
                        udo.LogTableName = def.LogTableName;
                }

                udo.EnableEnhancedForm = def.EnableEnhancedForm;
                udo.RebuildEnhancedForm = def.RebuildEnhancedForm;

                AddFindColumns(udo, def.FindColumns);

                EnsureDefaultFormColumns(def);

                AddChildTables(udo, def.ChildTables);

                AddFormColumns(udo, def.FormColumns);

                if (def.UseEnhancedFormColumns)
                {
                    AddEnhancedFormColumns(udo, def.EnhancedFormColumns);
                }

                AddMetadata(udo, "No se pudo registrar UDO " + def.Code + ".");
                _log.Info("UDO registrado: " + def.Code);
            }
            finally
            {
                ReleaseComObject(rs);
                ReleaseComObject(udo);
            }
        }

        private void AddFindColumns(UserObjectsMD udo, List<UdoFindColumn> cols)
        {
            if (cols == null || cols.Count == 0) return;

            for (int i = 0; i < cols.Count; i++)
            {
                var c = cols[i];
                if (c == null || string.IsNullOrWhiteSpace(c.Alias)) continue;

                udo.FindColumns.ColumnAlias = c.Alias;
                udo.FindColumns.ColumnDescription = string.IsNullOrWhiteSpace(c.Description) ? c.Alias : c.Description;

                if (i < cols.Count - 1)
                    udo.FindColumns.Add();
            }
        }

        private void AddChildTables(UserObjectsMD udo, List<UdoChildTable> childs)
        {
            if (childs == null || childs.Count == 0) return;

            for (int i = 0; i < childs.Count; i++)
            {
                var ct = childs[i];
                if (ct == null || string.IsNullOrWhiteSpace(ct.TableName)) continue;

                udo.ChildTables.TableName = ct.TableName; // SIN @
                if (i < childs.Count - 1)
                    udo.ChildTables.Add();
            }
        }

        private void AddFormColumns(UserObjectsMD udo, List<UdoFormColumn> cols)
        {
            if (cols == null || cols.Count == 0) return;

            for (int i = 0; i < cols.Count; i++)
            {
                var c = cols[i];
                if (c == null || string.IsNullOrWhiteSpace(c.Alias)) continue;

                udo.FormColumns.FormColumnAlias = c.Alias;
                udo.FormColumns.FormColumnDescription = string.IsNullOrWhiteSpace(c.Description) ? c.Alias : c.Description;
                udo.FormColumns.Editable = BoYesNoEnum.tYES;

                udo.FormColumns.SonNumber = c.SonNumber;

                if (i < cols.Count - 1)
                    udo.FormColumns.Add();
            }
        }

        private void AddEnhancedFormColumns(UserObjectsMD udo, List<UdoEnhancedFormColumn> cols)
        {
            if (cols == null || cols.Count == 0) return;

            for (int i = 0; i < cols.Count; i++)
            {
                var c = cols[i];
                if (c == null || string.IsNullOrWhiteSpace(c.Alias)) continue;

                udo.EnhancedFormColumns.ColumnAlias = c.Alias;
                udo.EnhancedFormColumns.ColumnDescription = string.IsNullOrWhiteSpace(c.Description) ? c.Alias : c.Description;

                udo.EnhancedFormColumns.Editable = BoYesNoEnum.tYES;
                udo.EnhancedFormColumns.ColumnIsUsed = BoYesNoEnum.tYES;

                udo.EnhancedFormColumns.ChildNumber = c.ChildNumber;     // 0 cabecera, 1..n hijas
                udo.EnhancedFormColumns.ColumnNumber = c.ColumnNumber;   // 1..n  (NO 0)

                if (i < cols.Count - 1)
                    udo.EnhancedFormColumns.Add();
            }
        }

        private void AddMetadata(dynamic businessObject, string errorMessage)
        {
            var result = businessObject.Add();
            if (result != 0)
            {
                //var company = GetCompany();
                _company.GetLastError(out int errorCode, out string errorDescription);
                throw new InvalidOperationException(errorMessage + " SAP(" + errorCode + "): " + errorDescription);
            }
        }

        private static void EnsureDefaultFormColumns(UdoDefinition def)
        {
            if (def.CanCreateDefaultForm != BoYesNoEnum.tYES) return;
            if (def.FormColumns.Count > 0) return;
            if (def.EnhancedFormColumns == null || def.EnhancedFormColumns.Count == 0) return;

            var groupedColumns = def.EnhancedFormColumns
                .Where(c => c != null && !string.IsNullOrWhiteSpace(c.Alias))
                .OrderBy(c => c.ChildNumber)
                .ThenBy(c => c.ColumnNumber)
                .ToList();

            foreach (var col in groupedColumns)
            {
                def.FormColumns.Add(new UdoFormColumn
                {
                    Alias = col.Alias,
                    Description = col.Description,
                    SonNumber = col.ChildNumber
                });
            }
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.FinalReleaseComObject(obj);
            }
        }
    }

}
