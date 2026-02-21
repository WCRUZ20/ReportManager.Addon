using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Core
{
    public sealed class SapApplication
    {
        public Application App { get; }

        private SapApplication(Application app)
        {
            App = app ?? throw new ArgumentNullException(nameof(app));
        }

        public static SapApplication Connect(string connectionStringFromArgs)
        {
            if (string.IsNullOrWhiteSpace(connectionStringFromArgs))
                throw new ArgumentException("No se recibió el connection string (args[0]).");

            var guiApi = new SboGuiApi();
            guiApi.Connect(connectionStringFromArgs);

            var app = guiApi.GetApplication(-1);
            return new SapApplication(app);
        }
    }
}
