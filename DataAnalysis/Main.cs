using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using System.Configuration;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace DataAnalysis
{
    public partial class Main : Form
    {
        string InputPath { get { return txtInputPath.Text; } }
        bool IsSameName { get { return checkSameName.Checked; } }
        string OutputPath { get { return txtOutputPath.Text; } }
        string SheetName { get { return txtSheetName.Text; } }
        bool IsFirstRowColumn { get { return checkColName.Checked; } }
        public int MaxRows { get { return Convert.ToUInt16(numRow.Value); } }
        public static Main main;

        public Main()
        {
            InitializeComponent();
            Text = Application.ProductName + " " + Application.ProductVersion;
            txtInputPath.Text = ConfigurationManager.AppSettings["inputPath"];
            txtOutputPath.Text = ConfigurationManager.AppSettings["outputPath"];
            txtSheetName.Text = ConfigurationManager.AppSettings["sheetName"];
            numRow.Value = Convert.ToUInt16(ConfigurationManager.AppSettings["numRow"]);
            main = this;
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                dynamic workExcel = null;
                //导入读取Excel
                DataTable dateExcel = NPOIHelper.Excel2DT(InputPath, SheetName, IsFirstRowColumn,ref workExcel);
                //按所需屏蔽符合要求的单元格
                DataTable dataExcelConceal = Data.Helper.changeDT(dateExcel);
                //显示
                dgv.DataSource = dataExcelConceal;
                //导出Excel
                //NPOIHelper.DT2Excel(dataExcelConceal, OutputPath, SheetName, true);
                NPOIHelper.DT2Excel_Cover(dataExcelConceal, workExcel, OutputPath, "DataAnalysis", true);
                MessageBox.Show("Done");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Excel:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            if (ofdInputPath.ShowDialog() == DialogResult.OK)
            {
                txtInputPath.Text = ofdInputPath.FileName;
            }
        }

        private void txtInputPath_TextChanged(object sender, EventArgs e)
        {
            if (IsSameName)
            {
                txtOutputPath.Text = txtInputPath.Text;
            }
        }

        private void checkSameName_CheckedChanged(object sender, EventArgs e)
        {
            if (IsSameName)
            {
                txtOutputPath.Text = txtInputPath.Text;
                txtOutputPath.ReadOnly = true;
            }
            else { txtOutputPath.ReadOnly = false; }
        }
    }
}
