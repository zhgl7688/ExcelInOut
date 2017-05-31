using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace EXCEL导入导出
{
    public class TableOperation
    {
        public static void createTable(DataTable dt, string connectionString,string tableName)
        {
            //如果目标表不存在则创建  
            string strSql = string.Format(" if object_id('{0}') is null create table {0}(", tableName);  //以sheetName为表名
            foreach (System.Data.DataColumn c in dt.Columns)
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
        }
        /// <summary>
        /// 动态增加数据表的列
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="tableName"></param>
        /// <param name="connectionString"></param>
        /// <param name="columnNames"></param>
        public static void AddTableColumn(string tableName, string connectionString, List<string> columnNames)
        {
            if (columnNames != null && columnNames.Count > 0)
            {
                //如果目标表不存在则创建  
                string strSql = string.Format(" alter table  {0} add  ", tableName);  //以sheetName为表名
                foreach (var ColumnName in columnNames)
                {
                    strSql += string.Format("[{0}] varchar(500) ,", ColumnName);
                }
                if (strSql.Contains("varchar(500)"))
                    strSql = strSql.Substring(0, strSql.Length - 2);

                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                }
            }

        }
        /// <summary>
        /// 比较两个表的列不同并返回dt2中不同列的集合
        /// </summary>
        /// <param name="dt1">需要导入的表</param>
        /// <param name="dt2">数据库中的表</param>
        /// <returns></returns>
        public static List<string> GetDiffColums(DataTable dt1, DataTable dt2)
        {
            List<string> sts = new List<string>();
            foreach (System.Data.DataColumn c in dt1.Columns)
            {
                if (!dt2.Columns.Contains(c.ColumnName))
                {
                    sts.Add(c.ColumnName);
                }
            }
            return sts;
        }
    }
}
