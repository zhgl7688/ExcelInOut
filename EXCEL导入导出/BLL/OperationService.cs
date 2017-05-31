using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Data.SqlClient;


namespace EXCEL导入导出.BLL
{
    public interface IOperationService
    {
        //      1、创建文件夹
        void CreateDirectory();
        //2、复制文件到创建文件夹中
        void CopyFile();
        //3、读取文件
        void ReadFile();
        //4、更新数据库
        void Updata();
        //5、删除文件
        void DelFile();

    }
    public class OperationService
    {
        const string dirName = "TempDir";
        private string newFileName;
        private string pathName;
        DataTable dt;
        private string tableName;
        private string connString;
        public OperationService(string pathName, string tableName, string connString)
        {
            this.pathName = pathName;
            this.tableName = tableName;
            this.connString = connString;
            CreateDirectory();
            CopyFile();
            ReadFile();
            DataCorrection();
            Updata();
            DelFile();
        }
        //      1、创建文件夹
        void CreateDirectory()
        {
            if (!Directory.Exists(dirName))
            {
                Directory.CreateDirectory(dirName);
            }
        }
        //2、复制文件到创建文件夹中
        void CopyFile()
        {
            newFileName = pathName.Substring(pathName.LastIndexOf('\\'));
            var newFile = dirName + @"\" + newFileName;
            if (File.Exists(newFile)) File.Delete(newFile);
            File.Copy(pathName, newFile);
        }
        //3、读取文件
        void ReadFile()
        {
            string path = dirName + @"\" + newFileName;
            string sheetName = ExcelOpertion.NewExcelSheetName(path)[0];
            dt = new ExcelHelper(path).ExcelToDataTable(sheetName, true);
            dt.TableName = sheetName.Replace("$", "").Replace("'", "");
            if (dt.Rows.Count > 1) dt.Rows.RemoveAt(0);

        }
        /// <summary>
        /// 数据校正
        /// </summary>
        void DataCorrection()
        {
            //取除多余的列
            string[] removeStr = new string[]{  "EP2润滑油脂当前环累计量","盾尾密封当前环累计量","电量环累计量",
              "注浆B液当前环累计量","注浆右上B液当前环累计量","注浆左下B液当前环累计量",
              "注浆左上B液当前环累计量","备份老环号"};
            for (int i = 0; i < removeStr.Length; i++)
            {
                dt.Columns.Remove(removeStr[i]);
            }
            //修改成所需要的列名
            string[] replaceName = new string[]{"Ring No",	"Record Date",	"Record Time",	"System Condition",	     "里程",	"俯仰角",	"TBM滚动角",	"前部水平位移",	"前部竖直位移",	"中部水平位移",
     "中部竖直位移",	"剩余里程",	"盾构水平倾度",	"盾构竖直倾度",	"刀盘工作压力",	"刀盘扭矩",	     "刀盘转速",	"表示刀盘正转",	"表示刀盘反转",	"齿轮油油温",	"主油箱油温",	"总推力",	
     "速度",	"压力泵压力",	"主推进油缸A行程",	"主推进油缸A压力",	"主推进油缸B行程",	     "主推进油缸B压力",	"主推进油缸C行程",	"主推进油缸C压力",	"主推进油缸D行程",	
     "主推进油缸D压力",	"铰接油缸A行程",	"铰接油缸B行程",	"铰接油缸C行程",	     "铰接油缸D行程",	"铰接油缸压力",	"1封前仓盾尾密封油压",	"2封前仓盾尾密封油压",	
     "3封前仓盾尾密封油压",	"4封前仓盾尾密封油压",	"1封后仓盾尾密封油压",	"2封后仓盾尾密封油压",     "3封后仓盾尾密封油压",	"4封后仓盾尾密封油压",	"土压1",	"土压2",	"土压3",	"土压4",	"土压5",	
     "螺旋输送机工作压力",	"螺旋输送机工作油温",	"螺旋输送机轴扭矩",	"螺旋输送机转速",	"螺旋输送机闸门开度",	"螺旋输送机土压",	"注浆量总和",	"1注浆压力",	"1注浆量",	"2注浆压力",	"2注浆量",
     "3注浆压力",	"3注浆量",	"4注浆压力",	"4注浆量",	"水泵流量",	"泡沫系统1压力",	"泡沫系统1空气流量",	"泡沫系统1添加剂流量",	"泡沫系统2压力",	"泡沫系统2空气流量",	"泡沫系统2添加剂流量",
     "泡沫系统3压力",	"泡沫系统3空气流量",	"泡沫系统3添加剂流量",	"泡沫系统4压力",	"泡沫系统4空气流量",	"泡沫系统4添加剂流量",	"设备桥压力",	"皮带机转速",	"工业水总累计量",	"注浆B液总累计量",
     "刀盘喷水总累计量",	"膨润土总累计量",	"盾壳膨润土总累计量",	"刀盘总累计工作时间",	"HBW密封油脂总累计量",	"EP2润滑油脂总累计量",	"电量总累计量",	"泡沫原液当前环累计量",	"工业水当前环累计量",	
     "泡沫混合液当前环累计量",	"刀盘喷水当前环累计量",	"膨润土当前环累计量",	"盾壳膨润土当前环累计量",	"刀盘当前累计工作时间",	"HBW密封油脂当前环累计量",	"EP2润滑油脂当前环累计量",	"盾尾密封当前环累计量",	
     "电量环累计量",	"注浆B液当前环累计量",	"注浆右上B液当前环累计量",	"注浆左下B液当前环累计量",	"注浆左上B液当前环累计量",	"备份老环号"
};
            for (int i = replaceName.Length - 1; i > 66; i--)
            {
                dt.Columns[i].ColumnName = replaceName[i];
            }


            //第二列日期校正
            dt.Rows[0][1] = dt.Rows[0][1].ToString().Split(' ')[0];
            //第三列时间校正
            dt.Rows[0][2] = dt.Rows[0][2].ToString().Split(' ')[1];
        }
        //4、更新数据库
        void Updata()
        {
            saveDataSet(dt);
        }
        //5、删除文件
        void DelFile()
        {
            File.Delete(dirName + @"\" + newFileName);
        }
        private void saveDataSet(DataTable dtEx)
        {
            //删除原有的列
            SQLHelper.Update("delete  from ods");

            using (System.Data.SqlClient.SqlConnection sqlconnt = new System.Data.SqlClient.SqlConnection(connString))
            {
                sqlconnt.Open();

                DataSet dst = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter("select top 1 * from ods where 1<>1", sqlconnt);
                SqlCommandBuilder cmd = new SqlCommandBuilder(ad);
                ad.Fill(dst, tableName);
                DataTable dtSourcet = dst.Tables[tableName];
                //创建没有的列表名
                List<string> diffcol = TableOperation.GetDiffColums(dtEx, dtSourcet);
                if (diffcol.Count > 0)//有不同的列
                {

                    TableOperation.AddTableColumn(tableName, connString, diffcol);
                }
            }

            SqlBulkCopyByDatatable(connString, tableName, dtEx);

        }
        private void SqlBulkCopyByDatatable(string connectionString, string TableName, DataTable dt)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction))
                {
                    try
                    {
                        sqlbulkcopy.DestinationTableName = TableName;
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sqlbulkcopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                        }
                        sqlbulkcopy.WriteToServer(dt);
                    }
                    catch (System.Exception ex)
                    {
                        throw ex;
                    }
                }
            }
        }
    }
}
