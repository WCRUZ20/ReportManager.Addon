using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Core
{
    public sealed class PrincipalFormController
    {
        private readonly Application _app;
        private readonly SrfFormLoader _loader;
        private readonly string _srfPath;
        private readonly string _formUid;

        public PrincipalFormController(Application app, SrfFormLoader loader, string srfPath, string formUid)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _loader = loader ?? throw new ArgumentNullException(nameof(loader));
            _srfPath = !string.IsNullOrWhiteSpace(srfPath) ? srfPath : throw new ArgumentException("srfPath es requerido.", nameof(srfPath));
            _formUid = !string.IsNullOrWhiteSpace(formUid) ? formUid : throw new ArgumentException("formUid es requerido.", nameof(formUid));
        }

        public void OpenOrFocus()
        {
            var form = TryGetOpenForm(_formUid);
            if (form != null)
            {
                form.Visible = true;
                form.Select();
                return;
            }

            _loader.LoadFromFile(_srfPath);
            var loadedForm = TryGetOpenForm(_formUid);
            if (loadedForm != null)
            {
                loadedForm.Visible = true;
                loadedForm.Select();
            }
        }

        private Form TryGetOpenForm(string formUid)
        {
            try
            {
                return _app.Forms.Item(formUid);
            }
            catch
            {
                return null;
            }
        }
    }

}
