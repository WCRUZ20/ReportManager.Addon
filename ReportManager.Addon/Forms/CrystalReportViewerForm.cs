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

namespace ReportManager.Addon.Forms
{
    public sealed partial class CrystalReportViewerForm : Form
    {
        //private readonly ReportDocument _reportDocument;
        //private readonly CrystalReportViewer _viewer;

        public CrystalReportViewerForm(ReportDocument reportDocument)
        {
            //_reportDocument = reportDocument ?? throw new ArgumentNullException(nameof(reportDocument));

            Text = "Visualizador Crystal Reports";
            Width = 1024;
            Height = 768;
            StartPosition = FormStartPosition.CenterScreen;

            //_viewer = new CrystalReportViewer
            //{
            //    Dock = DockStyle.Fill,
            //    ToolPanelView = ToolPanelViewType.None,
            //    ShowGroupTreeButton = true,
            //    ShowParameterPanelButton = true,
            //    ReuseParameterValuesOnRefresh = true,
            //    ReportSource = _reportDocument,
            //};

            //Controls.Add(_viewer);
            //_viewer.Show();
            //_viewer.Refresh();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            //_viewer.ReportSource = null;
            //_reportDocument.Close();
            //_reportDocument.Dispose();
            base.OnFormClosed(e);
        }
    }

}
