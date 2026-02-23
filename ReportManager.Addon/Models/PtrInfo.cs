using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Models
{
    internal sealed class PtrInfo
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; }
    }
}
