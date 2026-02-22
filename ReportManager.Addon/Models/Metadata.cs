using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Models
{
    public sealed class UdoDefinition
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string TableName { get; set; }          // SIN @
        public BoUDOObjType ObjectType { get; set; }   // MasterData / Document

        // Servicios principales
        public BoYesNoEnum CanFind { get; set; } = BoYesNoEnum.tYES;
        public BoYesNoEnum CanCancel { get; set; } = BoYesNoEnum.tNO;
        public BoYesNoEnum CanClose { get; set; } = BoYesNoEnum.tNO;
        public BoYesNoEnum CanDelete { get; set; } = BoYesNoEnum.tYES;
        public BoYesNoEnum CanCreateDefaultForm { get; set; } = BoYesNoEnum.tYES;
        public BoYesNoEnum CanYearTransfer { get; set; } = BoYesNoEnum.tNO;
        public BoYesNoEnum ManageSeries { get; set; } = BoYesNoEnum.tNO;

        // Log (opcional)
        public BoYesNoEnum UseLog { get; set; } = BoYesNoEnum.tNO;
        public string LogTableName { get; set; } // SIN @ (si aplica)

        public BoYesNoEnum EnableEnhancedForm { get; set; } = BoYesNoEnum.tYES;

        public BoYesNoEnum RebuildEnhancedForm { get; set; } = BoYesNoEnum.tYES;

        // Columnas de búsqueda (Find Form) :contentReference[oaicite:2]{index=2}
        public List<UdoFindColumn> FindColumns { get; } = new List<UdoFindColumn>();

        // Tablas hijas (SIN @)
        public List<UdoChildTable> ChildTables { get; } = new List<UdoChildTable>();

        // Columnas activas en el Default Form (Matrix style)
        // SonNumber: 0 = cabecera (tabla principal), 1..n = tabla hija # en orden de registro :contentReference[oaicite:3]{index=3}
        public List<UdoFormColumn> FormColumns { get; } = new List<UdoFormColumn>();

        // Si en algún punto deseas “Enhanced default form”, aquí lo dejamos listo
        public bool UseEnhancedFormColumns { get; set; } = false;
        public List<UdoEnhancedFormColumn> EnhancedFormColumns { get; } = new List<UdoEnhancedFormColumn>();
    }

    public sealed class UdoFindColumn
    {
        public string Alias { get; set; }          // Ej: "Code", "U_SS_IDRPT"
        public string Description { get; set; }    // Texto visible
    }

    public sealed class UdoChildTable
    {
        public string TableName { get; set; }      // SIN @. Ej: "SS_PRMDET"
    }

    public sealed class UdoFormColumn
    {
        public string Alias { get; set; }          // "Code", "Name", "U_MiCampo"
        public string Description { get; set; }    // Texto visible
        public int SonNumber { get; set; }         // 0 = principal; 1..n = tabla hija
    }

    public sealed class UdoEnhancedFormColumn
    {
        public string Alias { get; set; }
        public string Description { get; set; }
        public int ChildNumber { get; set; }       // 0 = principal; 1..n = hija :contentReference[oaicite:4]{index=4}
        public int ColumnNumber {  get; set; }
    }
}
