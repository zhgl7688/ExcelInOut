using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EXCEL导入导出
{
    public partial class FrmDelOfAdd : Form
    {
        string connString = ConfigurationManager.ConnectionStrings["connString"].ToString();
        List<string> sqlList;
        public FrmDelOfAdd()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("这个将删除库中所有数据再追加", "注意", MessageBoxButtons.OKCancel);
            if (result == DialogResult.Cancel) return;
            //1、打开文件
            OpenFileDialog file1 = new OpenFileDialog();
            file1.Filter = "excel文件|*.xls";//3、判断文件类型*.xlsx;
            if (file1.ShowDialog() == DialogResult.OK)  //2、判断文件是否存在
            {
                this.txtFileName.Text = file1.FileName;
            }
        }
        private void btnExecute_Click(object sender, EventArgs e)
        {
            string fileName = txtFileName.Text;
            if (fileName.Length > 0)
            {
                var sheetNames = ExcelSheetName(fileName);//获取第一个表名
                string sheetName = sheetNames[0].Substring(0, sheetNames[0].Length - 1);
                TransferData(fileName, sheetName, connString);//4、进行导入
            }
            else
            {
                MessageBox.Show("请选择Excel文件后，再导入");
            }

        }
        /// <summary>
        /// 将Excel文件导入数据库中，先删除库中表再导入
        /// </summary>
        /// <param name="excelFile"></param>
        /// <param name="sheetName"></param>
        /// <param name="connectionString"></param>
        public void TransferData(string excelFile, string sheetName, string connectionString)
        {
            DataSet ds = new DataSet();
            try
            {
                //获取全部数据  
                string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1;\";";
                // string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFile + ";" + "Extended Properties=Excel 8.0;";  
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}$]", sheetName);
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                myCommand.Fill(ds, sheetName);

                //如果目标表不存在则创建  
                string strSql = string.Format("if object_id('{0}') is not null drop table {0} if object_id('{0}') is null create table {0}(", sheetName);  //以sheetName为表名
                foreach (System.Data.DataColumn c in ds.Tables[0].Columns)
                {
                    strSql += string.Format("[{0}] varchar(255),", c.ColumnName);
                }
                strSql = strSql.Trim(',') + ")";

                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                }
                //用bcp导入数据  
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(connectionString))
                {
                    bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(bcp_SqlRowsCopied);
                    count = ds.Tables[0].Rows.Count;
                    bcp.BatchSize = count > 1000 ? ds.Tables[0].Rows.Count / 10 : 100; ;//每次传输的行数  
                    bcp.NotifyAfter = count > 100 ? ds.Tables[0].Rows.Count / 10 : 1;// 1000;//进度提示的行数  
                    bcp.DestinationTableName = sheetName;//目标表  
                    bcp.WriteToServer(ds.Tables[0]);
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        int count;
        public List<string> ExcelSheetName(string filepath)
        {
            List<string> al = new List<string>();
            string strConn;
            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filepath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1;\";";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable sheetNames = conn.GetOleDbSchemaTable
            (System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            conn.Close();
            foreach (DataRow dr in sheetNames.Rows)
            {
                al.Add(dr[2].ToString());
            }
            return al;
        }
        void bcp_SqlRowsCopied(object sender, System.Data.SqlClient.SqlRowsCopiedEventArgs e)
        {
            string record = e.RowsCopied.ToString();
            record = count - Convert.ToInt32(record) < 10 ? count.ToString() : record;
            //this.lblRecord.Text = string.Format("导入记录行数：{0},总记录数：{1}。", record, count);
            this.Update();
        }
        private void btnTB_Click(object sender, EventArgs e)
        {
            //1、打开文件
            OpenFileDialog file1 = new OpenFileDialog();
            file1.Filter = "excel文件|*.xls";//3、判断文件类型*.xlsx;
            if (file1.ShowDialog() == DialogResult.OK)  //2、判断文件是否存在
            {
                var sheetNames = ExcelSheetName(file1.FileName);//获取第一个表名
                string sheetName = sheetNames[0].Substring(0, sheetNames[0].Length - 1);
                AddData(file1.FileName, sheetName, connString, true);//4、进行导入
            }
        }
        public void AddData(string excelFile, string sheetName, string connectionString)
        {
            DataSet ds = new DataSet();
            try
            {
                //获取全部数据  
                string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1;\";";
                // string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFile + ";" + "Extended Properties=Excel 8.0;";  
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}$]", sheetName);
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                myCommand.Fill(ds, sheetName);

                DataTable dt = ds.Tables[sheetName];
                sqlList = new List<string>();

                //获取表名进行比较
                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    var sqlColnum = string.Format("select name from syscolumns where id=(select max(id) from sysobjects where xtype='u' and name='{0}')", sheetName);
                    SqlDataAdapter sqlCommand = new SqlDataAdapter(sqlColnum, connectionString);
                    sqlCommand.Fill(ds, "colums");

                }
                if (dt != null && dt.Rows.Count > 0)
                {
                    string sql = "insert into " + sheetName + "( ";
                    foreach (DataColumn item in dt.Columns)
                    {
                        bool sst = false;
                        foreach (DataRow sqlitem in ds.Tables["colums"].Rows)
                        {
                            if (sqlitem[0].ToString() == item.ColumnName)
                            {
                                sst = true;
                                break;
                            }
                        }
                        if (!sst)
                        {
                            MessageBox.Show("这个字段名不存在" + item.ColumnName);
                            return;
                        }
                        sql += item.ColumnName + ",";
                    }
                    sql = sql.Substring(0, sql.Length - 1);
                    sql += " ) values (";
                    //导入数据
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sqlt = sql;
                        foreach (DataColumn dc in dt.Columns)
                        {
                            var ss = dr[dc];
                            sqlt += "'" + dr[dc].ToString() + "',";
                        }
                        sqlt = sqlt.Substring(0, sqlt.Length - 1);
                        sqlt += " )";
                        sqlList.Add(sqlt);
                    }
                }

                var t = Task.Factory.StartNew(() => upData());


               // this.lblRecord.Text = string.Format("导入记录行数：{0}。", sqlList.Count);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        private void upData()
        {
            using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connString))
            {
                sqlconn.Open();
                System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                foreach (var strSql in sqlList)
                {
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                }

                sqlconn.Close();
            }
        }
        public void AddData(string excelFile, string sheetName, string connectionString, bool tb)
        {
            DataSet ds = new DataSet();

            //获取全部数据  
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1;\";";
            // string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFile + ";" + "Extended Properties=Excel 8.0;";  
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            strExcel = string.Format("select * from [{0}$]", sheetName);
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            myCommand.Fill(ds, sheetName);

            DataTable dt = ds.Tables[sheetName];


            //获取表名进行比较
            using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
            {
                sqlconn.Open();
                System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                var sqlColnum = string.Format("select name from syscolumns where id=(select max(id) from sysobjects where xtype='u' and name='{0}')", sheetName);
                SqlDataAdapter sqlCommand = new SqlDataAdapter(sqlColnum, connectionString);
                sqlCommand.Fill(ds, "colums");

            }
            if (dt != null && dt.Rows.Count > 0)
            {
                string sql = string.Format("insert into {0} (CompCode,PopCode,BrandCode,CateCode,DisplayNumber  ) values (@CompCode,@PopCode,@BrandCode,@CateCode,@DisplayNumber)", sheetName);
                List<SqlParameter[]> paraList = new List<SqlParameter[]>();
                //导入数据
                foreach (DataRow dr in dt.Rows)
                {
                    string sqlt = sql;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        var columnName = dc.ColumnName.ToString();
                        var columnNameS = columnName.Split(',');
                        if (columnNameS.Length > 1 && dr[dc].ToString() != "")
                        {
                            paraList.Add(new SqlParameter[]
                                         {
                                    new SqlParameter("@CompCode",dr["CompCode"]),
                                   new SqlParameter("@PopCode",dr["PopCode"]),
                                    new SqlParameter("@BrandCode",columnNameS[0]),
                                   new SqlParameter("@CateCode",columnNameS[1]),
                                   new SqlParameter("@DisplayNumber", dr[dc]),
                                           });
                        }
                    }
                }
                if (paraList.Count > 0) SQLHelper.Update(sql, paraList);
            }
        }

        
    }
}
