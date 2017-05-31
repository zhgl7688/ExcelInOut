using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EXCEL导入导出.BLL;

namespace EXCEL导入导出
{
    /// <summary>
    /// 功能：Excel导入时
    /// 1、数据库无表会自动创建表导入；
    /// 2、数据库有表有不一样的字段会自动创建导入；
    /// 3、先删除数据库有标识列与Excel表名相同的，再导入；
    /// </summary>
    public partial class FrmCur : Form
    {
        string connString = ConfigurationManager.ConnectionStrings["connString"].ToString();
        string tableName = ConfigurationManager.AppSettings["tableName"].ToString();
       
        string FilePath = ConfigurationManager.AppSettings["FilePath"].ToString();
        OperationService operationService;

       
        Timer timer = new Timer();
        public FrmCur()
        {
            InitializeComponent();
            timer.Interval = 10000;
            timer.Enabled = false;
            timer.Tick += load;

            lbUpData.Items.Add("功能：智能转换：将Excel数据转为Serve数据库数据");
             lbUpData.Items.Add("1、数据库有表但有不一样的字段会自动创建导入；");
            lbUpData.Items.Add("2、先删除数据库有标识列与Excel表名相同的，再导入；");
            lbUpData.Items.Add("使用方法:");
            lbUpData.Items.Add("1、修改appconfig中数据库连接字符串；");
            lbUpData.Items.Add("2、修改appconfig中存放Excel文件的文件夹；");
            lbUpData.Items.Add("3、新版本Excel导入需要安装AccessDatabaseEngine.exe；");
            lbUpData.Items.Add("4、数据库连接状态：");
            lbUpData.Items.Add(SQLHelper.SqlConnTest());
            lbUpData.Items.Add("——————*******—————————");
            this.button1.Visible = true;
           
        }
        int i = 0;
        //定时读取
        private void load(object sender, EventArgs e)
        {
            lbUpData.Items.Add("转换中"+i++.ToString());
            if (i%20== 0)
            {
                lbUpData.Items.Clear();
            }
            readFile();
        }
        //读出文件夹中文件
        private void readFile()
        {

            operationService = new OperationService(FilePath, tableName, connString);
             
        }
        private void btnAddInput_Click(object sender, EventArgs e)
        {
            readFile();
        }
 
        
        private void cbAutomatic_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
            {
                lbUpData.Items.Add("开启自动模式，请把Excel文件复到" + FilePath + "文件夹下");
                timer.Start();
            }
            else
            {
                lbUpData.Items.Add("关闭自动自动模式");
                timer.Stop();
            }
        }
    }
}
