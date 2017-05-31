using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EXCEL导入导出
{
    public class Log
    {
        public string file_name { get; set; }

        public string sheet_name { get; set; }
        public string ErrorPrompt { get; set; }

        public void saveLog()
        {
            string sql = string.Format("insert Import_info(file_name,sheet_name,ErrorPrompt) values('{0}','{1}','{2}')", file_name, sheet_name, ErrorPrompt);
            SQLHelper.Update(sql);
        }
    }
}
