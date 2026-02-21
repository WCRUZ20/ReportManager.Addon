using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Core
{
    public sealed class SrfFormLoader
    {
        private readonly Application _app;

        public SrfFormLoader(Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void LoadFromFile(string srfPath)
        {
            if (!File.Exists(srfPath))
                throw new FileNotFoundException("No se encontró el SRF en: " + srfPath);

            var xml = File.ReadAllText(srfPath);
            _app.LoadBatchActions(xml);

            // Si hubo errores en batch
            string result = _app.GetLastBatchResults();
            // result a veces viene vacío aunque cargue ok, pero si falla suele traer detalles.
        }
    }
}
