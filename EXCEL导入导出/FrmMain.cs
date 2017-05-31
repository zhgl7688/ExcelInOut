using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace EXCEL导入导出
{

    public partial class FrmMain : Form
    {
        Timer timer = new Timer();
        string connString = ConfigurationManager.ConnectionStrings["connString"].ToString();
        public FrmMain()
        {
            InitializeComponent();
            timer.Interval = 1000;
            timer.Enabled = false;
            timer.Tick += load;
         
            tSPBOUT.Width = this.Width - 100;
        }
        //定时读取
        private void load(object sender, EventArgs e)
        {
            readFile();
        }

        
        /// <summary>
        /// 获取数据库表名列表
        /// </summary>
        /// <param name="strConn"></param>
        /// <returns></returns>
        private List<string> GetTableList(string strConn)
        {
            List<string> al = new List<string>();
            SqlConnection conn = new SqlConnection(strConn);
            conn.Open();
            SqlDataAdapter myCommand = null;
            string strSQL = string.Format("Select Name FROM SysObjects Where XType='U' orDER BY Name");
            myCommand = new SqlDataAdapter(strSQL, strConn);
            DataSet ds = new DataSet();
            myCommand.Fill(ds);
            conn.Close();
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                al.Add(dr[0].ToString());
            }
            return al;
        }
        /// <summary>
        /// 获取数据库表数据
        /// </summary>
        /// <param name="strConn"></param>
        /// <returns></returns>
        private DataSet GetTableData(string strConn, string tableName)
        {
            SqlConnection conn = new SqlConnection(strConn);
            conn.Open();
            SqlDataAdapter myCommand = null;
            string strSQL = string.Format("Select * FROM  " + tableName);
            myCommand = new SqlDataAdapter(strSQL, strConn);
            DataSet ds = new DataSet();
            myCommand.Fill(ds);
            conn.Close();
            return ds;
        }
        /// <summary>
        /// 获取Excel表名
        /// </summary>
        /// <param name="filepath">Excel路径文件名</param>
        /// <returns></returns>
        

        
         

        private void btnExport_Click(object sender, EventArgs e)
        {
            //if (this.cboTableList.Text != null)
            //{
            //    SaveFileDialog file1 = new SaveFileDialog();
            //    file1.Filter = "Excel文件(*.xls)|*.xls";
            //    file1.FilterIndex = 1;
            //    file1.RestoreDirectory = true;

            //    if (file1.ShowDialog() == DialogResult.OK)
            //    {
            //        DataSet ds = GetTableData(connString, this.cboTableList.Text);
            //        DataChangeExcel datachange = new DataChangeExcel();
            //        datachange.DataSetToExcel(ds, file1.FileName, false);
            //    }

            //}
            //else
            //{
            //    MessageBox.Show("请选择要导出的表");
            //}
        }
       
        
       
        /// <summary>
        /// 多个Excel文件导入数据库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void btnMulInport_Click(object sender, EventArgs e)
        //{
        //    var result = MessageBox.Show("这个将删除库中所有数据再追加", "注意", MessageBoxButtons.OKCancel);
        //    if (result == DialogResult.Cancel) return;
        //    OpenFileDialog openFiles = new OpenFileDialog();
        //    openFiles.Multiselect = true;
        //    openFiles.Filter = "excel文件|*.xls";//3、判断文件类型*.xlsx;
        //    if (openFiles.ShowDialog() == DialogResult.OK)  //2、判断文件是否存在
        //    {
        //        string connstring = ConfigurationManager.ConnectionStrings["connString"].ToString();
        //        foreach (var item in openFiles.FileNames)
        //        {
        //            var sheetNames = ExcelSheetName(item);//获取第一个表名
        //            string sheetName = sheetNames[0].Substring(0, sheetNames[0].Length - 1);
        //            TransferData(item, sheetName, connstring);//4、进行导入
        //        }
        //    }
        //}
      
        private void btnAdd_Click(object sender, EventArgs e)
        {

        }

        
        /// <summary>  
        /// 修改config文件(ConnectionString节点)  
        /// </summary>  
        /// <param name="name">键</param>  
        /// <param name="value">要修改成的值</param>  
      

       

        private void btnTest_Click(object sender, EventArgs e)
        {

        }

        private void cboTableList_DropDown(object sender, EventArgs e)
        {
            // this.cboTableList.DataSource = GetTableList(connString);

        }
        /// <summary>
        /// 下载带进度条代码(状态栏式进度条）
        /// </summary>
        /// <param name="URL">网址</param>
        /// <param name="Filename">文件名</param>
        /// <param name="Prog">状态栏式进度条ToolStripProgressBar</param>
        /// <returns>True/False是否下载成功</returns>
        public bool DownLoadFile(string URL, string Filename, ToolStripProgressBar Prog)
        {
            try
            {
                System.Net.HttpWebRequest Myrq = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(URL); //从URL地址得到一个WEB请求   
                System.Net.HttpWebResponse myrp = (System.Net.HttpWebResponse)Myrq.GetResponse(); //从WEB请求得到WEB响应   
                long totalBytes = myrp.ContentLength; //从WEB响应得到总字节数   
                Prog.Maximum = (int)totalBytes; //从总字节数得到进度条的最大值   
                System.IO.Stream st = myrp.GetResponseStream(); //从WEB请求创建流（读）   
                System.IO.Stream so = new System.IO.FileStream(Filename, System.IO.FileMode.Create); //创建文件流（写）   
                long totalDownloadedByte = 0; //下载文件大小   
                byte[] by = new byte[1024];
                int osize = st.Read(by, 0, (int)by.Length); //读流   
                while (osize > 0)
                {
                    totalDownloadedByte = osize + totalDownloadedByte; //更新文件大小   
                    Application.DoEvents();
                    so.Write(by, 0, osize); //写流   
                    Prog.Value = (int)totalDownloadedByte; //更新进度条   
                    osize = st.Read(by, 0, (int)by.Length); //读流   
                }
                so.Close(); //关闭流   
                st.Close(); //关闭流   
                MessageBox.Show("导出成功！");
                Prog.Value = 0;
                return true;
            }
            catch
            {
                return false;
            }

        }

        private void btnOutTxt_Click(object sender, EventArgs e)
        {
            string URL = "http://192.168.1.166:8008/upload/1.txt";

            SaveFileDialog file1 = new SaveFileDialog();
            file1.Filter = "Txt文件(*.Txt)|*.Txt";
            file1.FilterIndex = 1;
            file1.RestoreDirectory = true;

            if (file1.ShowDialog() == DialogResult.OK)
            {
                DownLoadFile(URL, file1.FileName, this.tSPBOUT);

            }

        }

        private void btnOutBOM_Click(object sender, EventArgs e)
        {
            string URL = ConfigurationManager.AppSettings["BOMUri"].ToString();

            SaveFileDialog file1 = new SaveFileDialog();
            file1.Filter = "Txt文件(*.Txt)|*.Txt";
            file1.FilterIndex = 1;
            file1.RestoreDirectory = true;

            if (file1.ShowDialog() == DialogResult.OK)
            {
                DownLoadFile(URL, file1.FileName, this.tSPBOUT);

            }
        }

        private void btnOutCST_Click(object sender, EventArgs e)
        {
            string URL = ConfigurationManager.AppSettings["CSTUri"].ToString();
            SaveFileDialog file1 = new SaveFileDialog();
            file1.Filter = "Txt文件(*.Txt)|*.Txt";
            file1.FilterIndex = 1;
            file1.RestoreDirectory = true;

            if (file1.ShowDialog() == DialogResult.OK)
            {
                DownLoadFile(URL, file1.FileName, this.tSPBOUT);

            }
        }
        //读出文件夹中文件
        private void readFile()
        {
            string path = "EXCELInput";
            DirectoryInfo folder = new DirectoryInfo(path);
            foreach (FileInfo file in folder.GetFiles("*.xls"))
            {
                if (!File.Exists("alreadList.txt"))
                {
                    FileStream myFs = new FileStream("alreadList.txt", FileMode.Create);
                    StreamWriter mySw = new StreamWriter(myFs);
                    mySw.Close();
                    myFs.Close();
                }
             //判断是否调用过
                String[] lines = File.ReadAllLines("alreadList.txt");
                if (lines.Contains(file.FullName)) break;

                 
                DataSet dt = ExcelToTable(file.FullName);// Console.WriteLine(file.FullName);
                saveDataSet(dt);
                //保存调用文件
                using (StreamWriter sw = new StreamWriter("alreadList.txt", true))
                {
                    sw.WriteLine(file.FullName);
                }
            }

        }

 
        //dataset更新数据库
        private void saveDataSet(DataSet dsInput)
        {
            using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connString))
            {
                try
                {
                    sqlconn.Open();
                    DataSet ds = new DataSet();
                    foreach (DataTable dt in dsInput.Tables)
                    {
                        SqlDataAdapter ad = new SqlDataAdapter("select top 1 * from " + dt.TableName + " where 1<>1", sqlconn);

                        SqlCommandBuilder cmd = new SqlCommandBuilder(ad);

                        ad.Fill(ds, dt.TableName);
                        DataTable dtSource = ds.Tables[dt.TableName];
                        //合并table
                        {

                            foreach (DataRow item in dt.Rows)
                            {
                                object[] obj = new object[ds.Tables[dt.TableName].Columns.Count];
                                try
                                {
                                    item.ItemArray.CopyTo(obj, 0);
                                }
                                catch (Exception ex)
                                {
                                    lbUpData.Items.Add("表：" + dt.TableName + " 的列与数据库中表列不一致！");
                                    MessageBox.Show( dt.TableName+":"+ ex.Message);
                                    return;
                                }

                                dtSource.Rows.Add(obj);
                            }
                            ad.Update(ds, dt.TableName);
                            lbUpData.Items.Add("表名：" + dt.TableName + "更新记录数：" + dt.Rows.Count);
                        }




                    }
                }
                finally
                {
                    sqlconn.Close();
                }
            }

        }
        private void btnAddInput_Click(object sender, EventArgs e)
        {

           
           OpenFileDialog sflg = new OpenFileDialog();
           sflg.Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx";
           if (sflg.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
           }
           DataSet dt = ExcelToTable(sflg.FileName);// Console.WriteLine(file.FullName);
           saveDataSet(dt);
        }

        private static DataSet ExcelToTable(string fileName)
        {
            DataSet ds = new DataSet();
            DataTable dt = null;
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook(fs);
            int sheetCount = book.NumberOfSheets;
            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++)
            {
                NPOI.SS.UserModel.ISheet sheet = book.GetSheetAt(sheetIndex);
                if (sheet == null) continue;

                NPOI.SS.UserModel.IRow row = sheet.GetRow(0);
                if (row == null) continue;

                int firstCellNum = row.FirstCellNum;
                int lastCellNum = row.LastCellNum;
                if (firstCellNum == lastCellNum) continue;

                dt = new DataTable(sheet.SheetName);
                for (int i = firstCellNum; i < lastCellNum; i++)
                {
                    dt.Columns.Add(row.GetCell(i).StringCellValue, typeof(string));
                }

                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow newRow = dt.Rows.Add();
                    for (int j = firstCellNum; j < lastCellNum; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            sheet.GetRow(i).GetCell(j).SetCellType(NPOI.SS.UserModel.CellType.String);
                            newRow[j] = sheet.GetRow(i).GetCell(j).StringCellValue;
                        }

                    }
                }
                ds.Tables.Add(dt);
            }

            return ds;
        }

        private void cbAutomatic_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
            {
                lbUpData.Items.Add("开启自动导入EXCEL模式，请把Excel文件复到EXCELInput文件夹下");
                timer.Start();
            }else
            {
                lbUpData.Items.Add("关闭自动导入EXCEL模式");
                timer.Stop();
            }
        }
    }


}



