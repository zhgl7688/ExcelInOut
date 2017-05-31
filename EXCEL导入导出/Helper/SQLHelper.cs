using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace EXCEL导入导出
{
    /// <summary>
    /// 通用数据访问类
    /// </summary>
    public class SQLHelper
    {
        //数据库连接字符串
        private static string connString = ConfigurationManager.ConnectionStrings["connString"].ToString();
        #region 一般SQL语句数据访问
        public static string SqlConnTest()
        {
            using (SqlConnection conn = new SqlConnection(connString))
            {
                try
                {
                    conn.Open();
                    return "连接成功";
                }
                catch (SqlException err)
                {
                    return err.Message;
                }


            }
        }
        //数据增删改操作
        public static int Update(string sql)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                return cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

                throw new Exception(" 执行Update(string sql)时错误，具体信息：" + ex.Message+sql);
            }
            finally
            {
                conn.Close();
            }
        }
        //获取单个数据操作
        public static object GetSingleResult(string sql)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                return cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {

                throw new Exception(" 执行GetSingleResult (string sql)时错误，具体信息：" + ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        //获取只读数据集操作
        public static SqlDataReader GetReader(string sql)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                return cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                conn.Close();
                throw new Exception(" 执行GetReader(string sql)时错误，具体信息：" + ex.Message);
            }
        }
        //获取Dataset数据集操作
        public static DataSet GetDataSet(string sql)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            try
            {
                conn.Open();
                da.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw new Exception(" 执行GetDataSet(string sql)时错误，具体信息：" + ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        #endregion
        #region 带参数的SQL语句数据访问
        //数据增删改操作
        public static int Update(string sql,SqlParameter [] param)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                cmd.Parameters.AddRange(param);
                return cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

                throw new Exception(" 执行Update(string sql)时错误，具体信息：" + ex.Message+sql);
            }
            finally
            {
                conn.Close();
            }
        }
        //获取单个数据操作
        public static object GetSingleResult(string sql, SqlParameter[] param)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                cmd.Parameters.AddRange(param);
                return cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {

                throw new Exception(" 执行GetSingleResult (string sql)时错误，具体信息：" + ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        //获取只读数据集操作
        public static SqlDataReader GetReader(string sql, SqlParameter[] param)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                cmd.Parameters.AddRange(param);
                return cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                conn.Close();
                throw new Exception(" 执行GetReader(string sql)时错误，具体信息：" + ex.Message);
            }
        }
        //获取Dataset数据集操作
        public static DataSet GetDataSet(string sql, SqlParameter[] param)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            try
            {
                conn.Open();
                cmd.Parameters.AddRange(param);
                da.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw new Exception(" 执行GetDataSet(string sql)时错误，具体信息：" + ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        #endregion
        /// <summary>
        /// 带参数存储过程返回只读数据集
        /// </summary>
        /// <param name="proName">存储过程名称</param>
        /// <param name="param">参数集合</param>
        /// <returns></returns>
        public static SqlDataReader GetReaderByProc(string proName,SqlParameter[] param)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText =proName ;
            try
            {
                conn.Open();
                cmd.Parameters.AddRange(param);
               return  cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                conn.Close();
                throw new Exception(" 执行GetReaderByProc时错误，具体信息：" + ex.Message);
            }
        }
        /// <summary>
        /// 启用事务调用带参数的存储过程
        /// </summary>
        /// <param name="spName">存储过程名称</param>
        /// <param name="detailParam">存储参数</param>
        /// <returns></returns>
        public static bool Update(string sql, List<SqlParameter[]> Params)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            try
            {
                conn.Open();
                cmd.CommandText = sql;
                cmd.Transaction = conn.BeginTransaction();//开启事务
                foreach (SqlParameter[] param in Params)
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddRange(param);
                    cmd.ExecuteNonQuery();
                }
                cmd.Transaction.Commit();//提交事务
                return true;
            }
            catch (Exception ex)
            {
                if (cmd.Transaction != null)
                {
                    cmd.Transaction.Rollback();//回滚事务
                }

                //将异常信息写入日志 
                string errorInfo = "调用UpdateByTran(string mainSql,  SqlParameter[] mainParam,string detailSql ,List <SqlParameter []>detailParam)方法时发生错误，具体信息：" + ex.Message;
                 
                throw ex;
            }
            finally
            {
                if (cmd.Transaction != null)
                {
                    cmd.Transaction = null;//清空事务
                }
                conn.Close();
            }
        }
        //存储过程数据增删改操作
        public static int UpdateByProc(string procName, SqlParameter[] param)
        {
              SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = procName;
            try
            {
                conn.Open();
                cmd.Parameters.AddRange(param);
               return cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                string errorMsg ="存储过程：" +procName+" ----执行UpdateByProc(string procName, SqlParameter[] param)时错误，具体信息：" + ex.Message;
                new Log() { ErrorPrompt = errorMsg }.saveLog();
                throw new Exception(errorMsg);
            }
            finally
            {
                conn.Close();
            }
        }
        //存储过程数据增删改操作
        public static int UpdateByProc(string procName)
        {
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = procName;
            try
            {
                conn.Open(); 
                return cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                string errorMsg = "存储过程：" + procName + " ----执行UpdateByProc(string procName, SqlParameter[] param)时错误，具体信息：" + ex.Message;
                new Log() { ErrorPrompt = errorMsg }.saveLog();
                throw new Exception(errorMsg);
            }
            finally
            {
                conn.Close();
            }
        }
        /// <param name="table">准备更新的DataTable新数据</param>
        /// <param name="TableName">对应要更新的数据库表名</param>
        /// <param name="columnsName">对应要更新的列的列名集合</param>
        /// <param name="limitColumns">需要在ＳＱＬ的ＷＨＥＲＥ条件中限定的条件字符串，可为空。</param>
        /// <param name="onceUpdateNumber">每次往返处理的行数</param>
        /// <returns>返回更新的行数</returns>
        public static int Update(DataTable table, string TableName,  string[] columnsName, string limitWhere, int onceUpdateNumber)
        {
            if (string.IsNullOrEmpty(TableName)) return 0;
            if (columnsName == null || columnsName.Length <= 0) return 0;
            DataSet ds = new DataSet();
            ds.Tables.Add(table);
            int result = 0;
            using (SqlConnection sqlconn = new SqlConnection(connString))
            {
                sqlconn.Open();

                //使用加强读写锁事务   
                SqlTransaction tran = sqlconn.BeginTransaction(IsolationLevel.ReadCommitted);
                try
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        //所有行设为修改状态   
                        dr.SetModified();
                    }
                    //为Adapter定位目标表   
                    SqlCommand cmd = new SqlCommand(string.Format("select * from {0} where {1}", TableName, limitWhere), sqlconn, tran);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder(da);
                    da.AcceptChangesDuringUpdate = false;
                    string columnsUpdateSql = "";
                    SqlParameter[] paras = new SqlParameter[columnsName.Length];
                    //需要更新的列设置参数是,参数名为"@+列名"
                    for (int i = 0; i < columnsName.Length; i++)
                    {
                        //此处拼接要更新的列名及其参数值
                        columnsUpdateSql += ("[" + columnsName[i] + "]" + "=@" + columnsName[i] + ",");
                        paras[i] = new SqlParameter("@" + columnsName[i], columnsName[i]);
                    }
                    if (!string.IsNullOrEmpty(columnsUpdateSql))
                    {
                        //此处去掉拼接处最后一个","
                        columnsUpdateSql = columnsUpdateSql.Remove(columnsUpdateSql.Length - 1);
                    }

                    SqlCommand updateCmd = new SqlCommand(string.Format(" UPDATE [{0}] SET {1} ", TableName, columnsUpdateSql));
                    //不修改源DataTable   
                    updateCmd.UpdatedRowSource = UpdateRowSource.None;
                    da.UpdateCommand = updateCmd;
                    da.UpdateCommand.Parameters.AddRange(paras);
                    //每次往返处理的行数
                    da.UpdateBatchSize = onceUpdateNumber;
                    result = da.Update(ds, TableName);
                    ds.AcceptChanges();
                    tran.Commit();

                }
                catch
                {
                    tran.Rollback();
                }
                finally
                {
                    sqlconn.Dispose();
                    sqlconn.Close();
                }


            }
            return result;
        }
    }
}
