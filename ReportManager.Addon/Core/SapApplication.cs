using ReportManager.Addon.Entidades;
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

            var company = new SAPbobsCOM.Company();
            var contextCookie = company.GetContextCookie();
            var loginContext = app.Company.GetConnectionContext(contextCookie);

            var setContextResult = company.SetSboLoginContext(loginContext);
            if (setContextResult != 0)
                throw new InvalidOperationException($"Error SetSboLoginContext: {setContextResult}");

            var connectResult = company.Connect();
            if (connectResult != 0)
            {
                company.GetLastError(out int errorCode, out string errorMessage);
                throw new InvalidOperationException(
                    $"Error al conectar DI API. Código: {errorCode}. Mensaje: {errorMessage}"
                );
            }

            Globals.rCompany = company;
            return new SapApplication(app);
        }
    }
}
