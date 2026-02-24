using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Entidades
{
    class Globals
    {
        public static SAPbobsCOM.Company rCompany { get; set; }
        public static string dbuser { get; set; }
        public static string pwduser { get; set; }
        public static string chkrptSAP { get; set; }
    }
}
