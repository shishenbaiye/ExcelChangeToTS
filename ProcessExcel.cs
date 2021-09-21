using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;
using WpfApp;

namespace ExcelData
{
    class ProcessExcel
    {
        private static string _fileName;
        public static void Excel(string File) {
            DataTable dt = ReadExcelToTable(File);
            var count = dt.Rows.Count;
            var JsonString = new StringBuilder();
            if (count>0) {
                JsonString.Append("[");
                for (int i = 3; i < count; i++)
                {
                    JsonString.Append("{");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j == dt.Columns.Count - 1)
                        {
                            if(dt.Rows[1][j].ToString() == "string")
                                JsonString.Append($"\"{dt.Rows[2][j].ToString()}\":\"{dt.Rows[i][j].ToString()}\"");
                            if(dt.Rows[1][j].ToString() == "number")
                                JsonString.Append($"\"{dt.Rows[2][j].ToString()}\":{dt.Rows[i][j].ToString()}");
                        }
                        else 
                        {
                            if (dt.Rows[1][j].ToString() == "string")
                                JsonString.Append($"\"{dt.Rows[2][j].ToString()}\":\"{dt.Rows[i][j].ToString()}\",");
                            if (dt.Rows[1][j].ToString() == "number")
                                JsonString.Append($"\"{dt.Rows[2][j].ToString()}\":{dt.Rows[i][j].ToString()},");
                        }  
                    }
                    if (i == count - 1)
                    {
                        JsonString.Append("}");
                    }
                    else 
                    {
                        JsonString.Append("},");
                    }
                   
                }
                JsonString.Append("]");
            }
            ProcessExcel.ChangeToTypeScript(JsonString);
            
        }
        public static void ChangeToTypeScript(StringBuilder json) {
            var Json = json;
            var fileName = System.IO.Path.GetFileNameWithoutExtension(ProcessExcel._fileName);
            var path = System.IO.Path.GetDirectoryName(ProcessExcel._fileName);
            
            // 生成TS文件
            using FileStream ts = new FileStream($"{path}\\{fileName}.ts", FileMode.OpenOrCreate, FileAccess.Write);
            byte[] buffer = Encoding.Default.GetBytes($"const JSONSTRING = {Json.ToString()}");
            ts.Write(buffer, 0, buffer.Length);
            MessageBox.Show($"生成TS文件，位置在{path}");
        }

        public static DataTable ReadExcelToTable(string Path) {
            ProcessExcel._fileName = Path;
            try
            {
                string strConn;
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1;\"";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = "SELECT * FROM  [Sheet1$]";
                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, "Sheet1");
                OleConn.Close();

                return OleDsExcle.Tables["Sheet1"];
            }
            catch(Exception err)
            {
                MessageBox.Show("数据绑定Excel失败!失败原因：" + err.Message, "提示信息");
                return null;

            }
        }
    }
}
