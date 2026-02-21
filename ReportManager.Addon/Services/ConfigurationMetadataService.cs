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

        public void CreateReportConfigurationStructures()
        {
            CreateParameterStructures();
            CreateDefinitionStructures();
            _app.StatusBar.SetText("Estructuras maestras y UDOs validadas.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }

        private void CreateParameterStructures()
        {
            CreateUserTableIfNotExists("SS_PRMCAB", "Parametros Reporte Cab", BoUTBTableType.bott_MasterData);
            CreateUserFieldIfNotExists("@SS_PRMCAB", "SS_IDRPT", "Id Reporte", BoFieldTypes.db_Alpha, 50);
            CreateUserFieldIfNotExists("@SS_PRMCAB", "SS_NOMBRPT", "Nombre Reporte", BoFieldTypes.db_Alpha, 50);

            CreateUserTableIfNotExists("SS_PRMCABDET", "Parametros Reporte Det", BoUTBTableType.bott_MasterDataLines);
            CreateUserFieldIfNotExists("@SS_PRMCABDET", "SS_IDPARAM", "Id Parametro", BoFieldTypes.db_Alpha, 50);
            CreateUserFieldIfNotExists("@SS_PRMCABDET", "SS_DSCPARAM", "Desc Parametro", BoFieldTypes.db_Alpha, 50);
            CreateUserFieldIfNotExists("@SS_PRMCABDET", "SS_ACTIVO", "Activo", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "Y", "N");

            RegisterMasterDataUdoIfNotExists(
                "SS_PRMCAB",
                "SS_PRMCAB",
                "Parametrización de reportes",
                "SS_PRMCABDET",
                "U_SS_IDRPT");
        }

        private void CreateDefinitionStructures()
        {
            CreateUserTableIfNotExists("SS_DFRPTCAB", "Definicion Reporte Cab", BoUTBTableType.bott_MasterData);
            CreateUserFieldIfNotExists("@SS_DFRPTCAB", "SS_IDDPT", "Id Departamento", BoFieldTypes.db_Alpha, 50, BoFldSubTypes.st_None, null, null, "SS_DPTS");

            CreateUserTableIfNotExists("SS_DFRPTDET", "Definicion Reporte Det", BoUTBTableType.bott_MasterDataLines);
            CreateUserFieldIfNotExists("@SS_DFRPTDET", "SS_IDRPT", "Id Reporte", BoFieldTypes.db_Alpha, 50, BoFldSubTypes.st_None, null, null, "SS_PRMCAB");

            RegisterMasterDataUdoIfNotExists(
                "SS_DFRPTCAB",
                "SS_DFRPTCAB",
                "Definición de reportes",
                "SS_DFRPTDET",
                "U_SS_IDDPT");
        }

        private void  CreateUserTableIfNotExists(string tableName, string description, BoUTBTableType type)
        {
            UserTablesMD userTables = null;
            Recordset recordset = null;

            try
            {
                var company = GetCompany();
                recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery($"SELECT TOP 1 1 FROM OUTB WHERE TableName = '" + tableName + "'");

                if (!recordset.EoF)
                {
                    return;
                }

                userTables = (UserTablesMD)company.GetBusinessObject(BoObjectTypes.oUserTables);
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
                var company = GetCompany();
                recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery("SELECT TOP 1 1 FROM CUFD WHERE TableID = '" + tableName + "' AND AliasID = '" + fieldName + "'");

                if (!recordset.EoF)
                {
                    return;
                }

                fields = (UserFieldsMD)company.GetBusinessObject(BoObjectTypes.oUserFields);
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

        private void RegisterMasterDataUdoIfNotExists(string code, string tableName, string name, string childTableName, string referenceFieldAlias)
        {
            UserObjectsMD udo = null;
            Recordset recordset = null;

            try
            {
                var company = GetCompany();
                recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery("SELECT TOP 1 1 FROM OUDO WHERE Code = '" + code + "'");

                if (!recordset.EoF)
                {
                    return;
                }

                udo = (UserObjectsMD)company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                udo.Code = code;
                udo.Name = name;
                udo.ObjectType = BoUDOObjType.boud_MasterData;
                udo.TableName = tableName;
                udo.CanFind = BoYesNoEnum.tYES;
                udo.CanCancel = BoYesNoEnum.tNO;
                udo.CanClose = BoYesNoEnum.tNO;
                udo.CanDelete = BoYesNoEnum.tYES;
                udo.CanCreateDefaultForm = BoYesNoEnum.tYES;
                udo.CanYearTransfer = BoYesNoEnum.tNO;
                udo.ManageSeries = BoYesNoEnum.tNO;

                udo.FindColumns.ColumnAlias = "Code";
                udo.FindColumns.ColumnDescription = "Código";
                udo.FindColumns.Add();
                udo.FindColumns.ColumnAlias = referenceFieldAlias;
                udo.FindColumns.ColumnDescription = "Referencia";

                udo.ChildTables.TableName = childTableName;

                AddMetadata(udo, "No se pudo registrar UDO " + code + ".");
                _log.Info("UDO registrado: " + code);
            }
            finally
            {
                ReleaseComObject(recordset);
                ReleaseComObject(udo);
            }
        }

        private SAPbobsCOM.Company GetCompany()
        {
            return (SAPbobsCOM.Company)_app.Company.GetDICompany();
        }

        private void AddMetadata(dynamic businessObject, string errorMessage)
        {
            var result = businessObject.Add();
            if (result != 0)
            {
                var company = GetCompany();
                company.GetLastError(out int errorCode, out string errorDescription);
                throw new InvalidOperationException(errorMessage + " SAP(" + errorCode + "): " + errorDescription);
            }
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
        }
    }

}
