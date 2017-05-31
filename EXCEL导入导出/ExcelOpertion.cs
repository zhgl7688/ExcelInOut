using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace EXCEL导入导出
{

    public class ExcelOpertion
    {
        /// <summary>
        /// 获取Excel表名
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static List<string> NewExcelSheetName(string filepath)
        {
            List<string> al = new List<string>();
            string strConn = string.Empty;
            if (filepath.ToLower().IndexOf(".xlsx") > 0) // 2007版本
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1;\";";
            else if (filepath.ToLower().IndexOf(".xls") > 0) // 2003版本
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filepath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1;\";";
            if (strConn == null) return al;
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable sheetNames = conn.GetOleDbSchemaTable
            (System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            conn.Close();
            foreach (DataRow dr in sheetNames.Rows)
            {
                if (dr[2].ToString() == "Col") continue;
                if (!dr[2].ToString().Contains("ilterDatabase"))
                {
                    al.Add(dr[2].ToString());
                    break;
                }

            }
            return al;
        }
        public static List<string> OldExcelSheetName(string filepath)
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
    }
}
