namespace EXCEL导入导出
{
    partial class FrmConn
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
            this.txtConstring = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnTest = new System.Windows.Forms.Button();
            this.btnUpdateConnString = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtConstring
            // 
            this.txtConstring.Location = new System.Drawing.Point(14, 50);
            this.txtConstring.Name = "txtConstring";
            this.txtConstring.Size = new System.Drawing.Size(361, 21);
            this.txtConstring.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "数据库连接字符串";
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(158, 96);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(87, 23);
            this.btnTest.TabIndex = 5;
            this.btnTest.Text = "测试";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // btnUpdateConnString
            // 
            this.btnUpdateConnString.Location = new System.Drawing.Point(259, 96);
            this.btnUpdateConnString.Name = "btnUpdateConnString";
            this.btnUpdateConnString.Size = new System.Drawing.Size(87, 23);
            this.btnUpdateConnString.TabIndex = 6;
            this.btnUpdateConnString.Text = "更新库连接";
            this.btnUpdateConnString.UseVisualStyleBackColor = true;
            // 
            // FrmConn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(459, 157);
            this.Controls.Add(this.txtConstring);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.btnUpdateConnString);
            this.Name = "FrmConn";
            this.Text = "数据库连接字符串管理";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtConstring;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Button btnUpdateConnString;
    }
}