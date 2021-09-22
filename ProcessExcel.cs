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
        private static DataTable data;
        public static void Excel(string File) {
            DataTable dt = ReadExcelToTable(File);
            ProcessExcel.data = dt;
            var count = dt.Rows.Count;
            var JsonString = new StringBuilder();
            var functionString = new StringBuilder();
            var methodType = new StringBuilder();
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
                                JsonString.Append($"\"{dt.Rows[0][j].ToString()}\":\"{dt.Rows[i][j].ToString()}\"");
                            if(dt.Rows[1][j].ToString() == "number")
                                JsonString.Append($"\"{dt.Rows[0][j].ToString()}\":{dt.Rows[i][j].ToString()}");
                        }
                        else 
                        {
                            if (dt.Rows[1][j].ToString() == "string")
                                JsonString.Append($"\"{dt.Rows[0][j].ToString()}\":\"{dt.Rows[i][j].ToString()}\",");
                            if (dt.Rows[1][j].ToString() == "number")
                                JsonString.Append($"\"{dt.Rows[0][j].ToString()}\":{dt.Rows[i][j].ToString()},");
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
                JsonString.Append("]\n");
                functionString.Append("[");
                methodType.Append("[");
                for (int k = 0; k < dt.Columns.Count; k++) {
                    if (k == dt.Columns.Count - 1)
                    {
                        functionString.Append($"\"{dt.Rows[0][k]}\"");
                        methodType.Append($"\"{dt.Rows[1][k]}\"");
                    }
                    else
                    {
                        functionString.Append($"\"{dt.Rows[0][k]}\",");
                        methodType.Append($"\"{dt.Rows[1][k]}\",");
                    }
                }
                functionString.Append("]\n");
            }
            ProcessExcel.ChangeToTypeScript(JsonString,functionString,methodType);
            
        }
        public static void ChangeToTypeScript(StringBuilder json,StringBuilder funName, StringBuilder type) {
            var Json = json;
            var FunName = funName;
            var MethodType = type;
            var fileName = System.IO.Path.GetFileNameWithoutExtension(ProcessExcel._fileName);
            var path = System.IO.Path.GetDirectoryName(ProcessExcel._fileName);

            var str = new StringBuilder();
            for (int i = 0; i < ProcessExcel.data.Columns.Count; i++) {
                str.Append($"/**{ProcessExcel.data.Rows[2][i]} */ \n  {ProcessExcel.data.Rows[0][i]}:{ProcessExcel.data.Rows[1][i]} \n");
            }
            var str2 = new StringBuilder();
            str2.Append($"static GetDataById(id:number):I{fileName}"+"{ \n" +
                "let Data:any = null; \n let "+"i"+fileName+":I"+fileName+"; \n"+
                $"JSONSTRING{fileName}.forEach((item,index)=>"+"{"+
                "if(item.ID == id) \n"+"Data = item"+
                "}"+") \n"+
                $"i{fileName} = Data as I{fileName} \n return i{fileName}");
            // 生成TS文件
            using FileStream ts = new FileStream($"{path}\\{fileName}.ts", FileMode.OpenOrCreate, FileAccess.Write);
            byte[] buffer = Encoding.Default.GetBytes($"const JSONSTRING{fileName} = {Json.ToString()} \n" +
                $"const Params{fileName}:Array<string> = {FunName.ToString()} \n" +
                "interface "+"I"+fileName+"{ \n \n"+ str.ToString() +"}"+
                "export class "+fileName+"{ \n \n"+ str2.ToString()+"}"+"}");
            ts.Write(buffer, 0, buffer.Length);
            MessageBox.Show($"生成TS文件，位置在{path}");
        }

        public static DataTable ReadExcelToTable(string Path) {
            ProcessExcel._fileName = Path;
            var sheetName = System.IO.Path.GetFileNameWithoutExtension(Path);
            try
            {
                string strConn;
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1;\"";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = $"SELECT * FROM  [{sheetName}$]";
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
