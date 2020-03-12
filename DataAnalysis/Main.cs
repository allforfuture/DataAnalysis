using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataAnalysis
{
    public partial class Main : Form
    {
        string filepath { get { return txtFilePath.Text; } }
        string sheetName { get { return txtSheetName.Text; } }

        bool isFirstRowColumn { get { return checkColName.Checked; } }

        public Main()
        {
            InitializeComponent();
            Text = Application.ProductName + " " + Application.ProductVersion;
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            DataTable dt=new NPOIHelper().ExcelToDataTable(filepath, sheetName, isFirstRowColumn);
            //大于1200小于301
            dgv.DataSource = dt;
        }
    }
}
