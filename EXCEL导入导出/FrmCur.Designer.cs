namespace EXCEL导入导出
{
    partial class FrmCur
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCur));
            this.lbUpData = new System.Windows.Forms.ListBox();
            this.cbAutomatic = new System.Windows.Forms.CheckBox();
            this.tTp = new System.Windows.Forms.ToolTip(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbUpData
            // 
            this.lbUpData.FormattingEnabled = true;
            resources.ApplyResources(this.lbUpData, "lbUpData");
            this.lbUpData.Name = "lbUpData";
            // 
            // cbAutomatic
            // 
            resources.ApplyResources(this.cbAutomatic, "cbAutomatic");
            this.cbAutomatic.Name = "cbAutomatic";
            this.cbAutomatic.UseVisualStyleBackColor = true;
            this.cbAutomatic.CheckedChanged += new System.EventHandler(this.cbAutomatic_CheckedChanged);
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnAddInput_Click);
            // 
            // FrmCur
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lbUpData);
            this.Controls.Add(this.cbAutomatic);
            this.Name = "FrmCur";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lbUpData;
        private System.Windows.Forms.CheckBox cbAutomatic;
        private System.Windows.Forms.ToolTip tTp;
        private System.Windows.Forms.Button button1;
    }
}