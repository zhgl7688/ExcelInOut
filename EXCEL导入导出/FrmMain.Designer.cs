namespace EXCEL导入导出
{
    partial class FrmMain
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.btnOutBOM = new System.Windows.Forms.Button();
            this.btnOutCST = new System.Windows.Forms.Button();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.tSPBOUT = new System.Windows.Forms.ToolStripProgressBar();
            this.btnAddInput = new System.Windows.Forms.Button();
            this.cbAutomatic = new System.Windows.Forms.CheckBox();
            this.lbUpData = new System.Windows.Forms.ListBox();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOutBOM
            // 
            this.btnOutBOM.Location = new System.Drawing.Point(321, 252);
            this.btnOutBOM.Name = "btnOutBOM";
            this.btnOutBOM.Size = new System.Drawing.Size(75, 23);
            this.btnOutBOM.TabIndex = 5;
            this.btnOutBOM.Text = "导出BOM";
            this.btnOutBOM.UseVisualStyleBackColor = true;
            this.btnOutBOM.Click += new System.EventHandler(this.btnOutBOM_Click);
            // 
            // btnOutCST
            // 
            this.btnOutCST.Location = new System.Drawing.Point(411, 252);
            this.btnOutCST.Name = "btnOutCST";
            this.btnOutCST.Size = new System.Drawing.Size(75, 23);
            this.btnOutCST.TabIndex = 6;
            this.btnOutCST.Text = "导出CST";
            this.btnOutCST.UseVisualStyleBackColor = true;
            this.btnOutCST.Click += new System.EventHandler(this.btnOutCST_Click);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel1,
            this.tSPBOUT});
            this.toolStrip1.Location = new System.Drawing.Point(0, 286);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(516, 25);
            this.toolStrip1.TabIndex = 8;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(56, 22);
            this.toolStripLabel1.Text = "导出进度";
            // 
            // tSPBOUT
            // 
            this.tSPBOUT.AutoSize = false;
            this.tSPBOUT.Name = "tSPBOUT";
            this.tSPBOUT.Size = new System.Drawing.Size(100, 22);
            // 
            // btnAddInput
            // 
            this.btnAddInput.Location = new System.Drawing.Point(411, 81);
            this.btnAddInput.Name = "btnAddInput";
            this.btnAddInput.Size = new System.Drawing.Size(93, 23);
            this.btnAddInput.TabIndex = 12;
            this.btnAddInput.Text = "选择文件导入";
            this.btnAddInput.UseVisualStyleBackColor = true;
            this.btnAddInput.Click += new System.EventHandler(this.btnAddInput_Click);
            // 
            // cbAutomatic
            // 
            this.cbAutomatic.AutoSize = true;
            this.cbAutomatic.Location = new System.Drawing.Point(411, 40);
            this.cbAutomatic.Name = "cbAutomatic";
            this.cbAutomatic.Size = new System.Drawing.Size(96, 16);
            this.cbAutomatic.TabIndex = 14;
            this.cbAutomatic.Text = "开启自动模式";
            this.cbAutomatic.UseVisualStyleBackColor = true;
            this.cbAutomatic.CheckedChanged += new System.EventHandler(this.cbAutomatic_CheckedChanged);
            // 
            // lbUpData
            // 
            this.lbUpData.FormattingEnabled = true;
            this.lbUpData.ItemHeight = 12;
            this.lbUpData.Location = new System.Drawing.Point(12, 12);
            this.lbUpData.Name = "lbUpData";
            this.lbUpData.Size = new System.Drawing.Size(384, 196);
            this.lbUpData.TabIndex = 15;
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(516, 311);
            this.Controls.Add(this.lbUpData);
            this.Controls.Add(this.cbAutomatic);
            this.Controls.Add(this.btnAddInput);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.btnOutBOM);
            this.Controls.Add(this.btnOutCST);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel导入导出数据库";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOutBOM;
        private System.Windows.Forms.Button btnOutCST;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripProgressBar tSPBOUT;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.Button btnAddInput;
        private System.Windows.Forms.CheckBox cbAutomatic;
        private System.Windows.Forms.ListBox lbUpData;
    }
}

