using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Windows.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportManager.Addon
{
    public partial class Form1 : Form
    {
        private readonly ReportDocument _reportDocument;
        private readonly CrystalReportViewer _viewer;

        public Form1(ReportDocument reportDocument)
        {
            InitializeComponent();

            _viewer = new CrystalReportViewer
            {
                Dock = DockStyle.Fill,
                ToolPanelView = ToolPanelViewType.None,
                ShowGroupTreeButton = true,
                ShowParameterPanelButton = true,
                ReuseParameterValuesOnRefresh = true,
            };

            Controls.Add(_viewer);
            _viewer.Show();
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
