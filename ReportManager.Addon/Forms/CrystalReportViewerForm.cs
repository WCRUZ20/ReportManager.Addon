using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Windows.Forms;
using System;
using System.Windows.Forms;

namespace ReportManager.Addon.Forms
{
    public sealed partial class CrystalReportViewerForm : Form
    {
        private readonly ReportDocument _reportDocument;
        private readonly CrystalReportViewer _viewer;

        public CrystalReportViewerForm(ReportDocument reportDocument)
        {
            _reportDocument = reportDocument ?? throw new ArgumentNullException(nameof(reportDocument));

            InitializeComponent();

            Text = "Visualizador Crystal Reports";
            Width = 1024;
            Height = 768;
            StartPosition = FormStartPosition.CenterScreen;

            _viewer = new CrystalReportViewer
            {
                Dock = DockStyle.Fill,
                ToolPanelView = ToolPanelViewType.None,
                ShowGroupTreeButton = true,
                ShowParameterPanelButton = true,
                ReuseParameterValuesOnRefresh = true,
            };

            Controls.Add(_viewer);
            Shown += OnViewerFormShown;
        }

        private void OnViewerFormShown(object sender, EventArgs e)
        {
            // Diferimos el enlace del ReportSource para evitar que el constructor
            // ejecute lógica pesada y bloquee la creación inicial de la ventana.
            BeginInvoke((MethodInvoker)(() =>
            {
                Cursor = Cursors.WaitCursor;
                try
                {
                    _viewer.ReportSource = _reportDocument;
                    _viewer.Refresh();
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }));
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _viewer.ReportSource = null;
            _reportDocument.Close();
            _reportDocument.Dispose();
            base.OnFormClosed(e);
        }
    }
}
