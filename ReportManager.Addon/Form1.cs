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

        public Form1()
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
        }
    }
}
