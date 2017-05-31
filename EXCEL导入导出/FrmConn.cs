using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace EXCEL导入导出
{
    public partial class FrmConn : Form
    {
        string connString = ConfigurationManager.ConnectionStrings["connString"].ToString();
        public FrmConn()
        {
            InitializeComponent();
            this.txtConstring.Text = connString;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(this.txtConstring.Text.Trim()))
            {
                try
                {
                    sqlconn.Open();
                    if (sqlconn.State == ConnectionState.Open)
                        MessageBox.Show("连接成功");

                }
                finally
                {
                    sqlconn.Close();
                }
            }
        }
        private void btnUpdateConnString_Click(object sender, EventArgs e)
        {
            UpdateConnectionString("connString", this.txtConstring.Text.Trim());
            MessageBox.Show("请重打开程序");
            Close();
        }
        public static void UpdateConnectionString(string name, string value)
        {
            XmlDocument doc = new XmlDocument();
            //获得配置文件的全路径   
            string strFileName = AppDomain.CurrentDomain.BaseDirectory.ToString() + "EXCEL导入导出.exe.config";
            doc.Load(strFileName);
            //找出名称为“add”的所有元素   
            XmlNodeList nodes = doc.GetElementsByTagName("add");
            for (int i = 0; i < nodes.Count; i++)
            {
                //获得将当前元素的key属性   
                XmlAttribute _name = nodes[i].Attributes["name"];
                //根据元素的第一个属性来判断当前的元素是不是目标元素   
                if (_name != null)
                {
                    if (_name.Value == name)
                    {
                        //对目标元素中的第二个属性赋值   
                        _name = nodes[i].Attributes["connectionString"];

                        _name.Value = value;
                        break;
                    }
                }
            }
            //保存上面的修改   
            doc.Save(strFileName);
        }

    }
}
