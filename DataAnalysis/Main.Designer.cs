namespace DataAnalysis
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.checkSameName = new System.Windows.Forms.CheckBox();
            this.pic = new System.Windows.Forms.PictureBox();
            this.ofdInputPath = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtInputPath = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.txtSheetName = new System.Windows.Forms.TextBox();
            this.checkColName = new System.Windows.Forms.CheckBox();
            this.numRow = new System.Windows.Forms.NumericUpDown();
            this.txtOutputPath = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRow)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(504, 24);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(75, 23);
            this.btnSelectFile.TabIndex = 9;
            this.btnSelectFile.Text = "选择文件";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(94, 7);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(173, 12);
            this.label5.TabIndex = 19;
            this.label5.Text = "操作前请先备份，数据仅供参考";
            // 
            // checkSameName
            // 
            this.checkSameName.AutoSize = true;
            this.checkSameName.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.checkSameName.Location = new System.Drawing.Point(12, 51);
            this.checkSameName.Name = "checkSameName";
            this.checkSameName.Size = new System.Drawing.Size(138, 16);
            this.checkSameName.TabIndex = 20;
            this.checkSameName.Text = "是否导出到原文件： ";
            this.checkSameName.UseVisualStyleBackColor = true;
            this.checkSameName.CheckedChanged += new System.EventHandler(this.checkSameName_CheckedChanged);
            // 
            // pic
            // 
            this.pic.Image = ((System.Drawing.Image)(resources.GetObject("pic.Image")));
            this.pic.Location = new System.Drawing.Point(585, 0);
            this.pic.Name = "pic";
            this.pic.Size = new System.Drawing.Size(344, 205);
            this.pic.TabIndex = 21;
            this.pic.TabStop = false;
            // 
            // ofdInputPath
            // 
            this.ofdInputPath.FileName = "openFileDialog1";
            this.ofdInputPath.Filter = "2003版本Excel文件|*.xls|2007版本Excel文件|*.xlsx";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "导入Excel文件：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 99);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "数据源工作表名称：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(49, 147);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "保留行数：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(34, 77);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(95, 12);
            this.label4.TabIndex = 3;
            this.label4.Text = "导出Excel文件：";
            // 
            // txtInputPath
            // 
            this.txtInputPath.Location = new System.Drawing.Point(135, 26);
            this.txtInputPath.Name = "txtInputPath";
            this.txtInputPath.Size = new System.Drawing.Size(354, 21);
            this.txtInputPath.TabIndex = 10;
            this.txtInputPath.TextChanged += new System.EventHandler(this.txtInputPath_TextChanged);
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(504, 72);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 23);
            this.btnRun.TabIndex = 11;
            this.btnRun.Text = "运行";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // dgv
            // 
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(12, 211);
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.Size = new System.Drawing.Size(906, 274);
            this.dgv.TabIndex = 12;
            // 
            // txtSheetName
            // 
            this.txtSheetName.Location = new System.Drawing.Point(135, 96);
            this.txtSheetName.Name = "txtSheetName";
            this.txtSheetName.Size = new System.Drawing.Size(100, 21);
            this.txtSheetName.TabIndex = 13;
            // 
            // checkColName
            // 
            this.checkColName.AutoSize = true;
            this.checkColName.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.checkColName.Checked = true;
            this.checkColName.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkColName.Enabled = false;
            this.checkColName.Location = new System.Drawing.Point(12, 121);
            this.checkColName.Name = "checkColName";
            this.checkColName.Size = new System.Drawing.Size(138, 16);
            this.checkColName.TabIndex = 14;
            this.checkColName.Text = "第一行是否是列名： ";
            this.checkColName.UseVisualStyleBackColor = true;
            // 
            // numRow
            // 
            this.numRow.Location = new System.Drawing.Point(135, 144);
            this.numRow.Name = "numRow";
            this.numRow.Size = new System.Drawing.Size(40, 21);
            this.numRow.TabIndex = 15;
            this.numRow.Value = new decimal(new int[] {
            87,
            0,
            0,
            0});
            // 
            // txtOutputPath
            // 
            this.txtOutputPath.Location = new System.Drawing.Point(135, 71);
            this.txtOutputPath.Name = "txtOutputPath";
            this.txtOutputPath.Size = new System.Drawing.Size(354, 21);
            this.txtOutputPath.TabIndex = 16;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(930, 497);
            this.Controls.Add(this.pic);
            this.Controls.Add(this.checkSameName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.txtOutputPath);
            this.Controls.Add(this.numRow);
            this.Controls.Add(this.checkColName);
            this.Controls.Add(this.txtSheetName);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.txtInputPath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Main";
            this.Text = "Form1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.pic)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox checkSameName;
        private System.Windows.Forms.PictureBox pic;
        private System.Windows.Forms.OpenFileDialog ofdInputPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtInputPath;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.TextBox txtSheetName;
        private System.Windows.Forms.CheckBox checkColName;
        private System.Windows.Forms.NumericUpDown numRow;
        private System.Windows.Forms.TextBox txtOutputPath;
    }
}

