using ExcelProtobuf;
using Microsoft.Win32.SafeHandles;
using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelProtobuf
{
    public class ProtoHandler
    {
        public void Process(string groupName,bool force)
        {
            Program.Log("***开始生成proto文件***");

            if(Config.instance.GroupHashCheck(groupName))
            {
                if (!force)
                {
                    Program.Log("组 {0} 无需更新", groupName);
                    throw new Exception();
                }
            }

            var mappings = Config.instance.GetMappings(groupName);

            var path = Config.instance.protoDirectory;
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
            Directory.CreateDirectory(path);

            for (int i = 0; i < mappings.Length; i++)
            {
                try
                {
                    ProcessExcel(mappings[i].datPath, mappings[i].excelName,groupName);
                }
                catch (Exception e)
                {
                    Config.instance.RecordLog(e);
                    Program.Log("生成失败 {0}", mappings[i].datPath);
                    continue;
                }
                Program.Log("生成成功 {0}", mappings[i].datPath);
            }
        }

        private void ProcessExcel(string datPath, string excelName,string groupName)
        {
            XSSFWorkbook workbook = null;
            using (FileStream fs = new FileStream(Config.instance.excelDirectory + excelName, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fs);
            }

            try
            {
                string protoName = Path.GetFileNameWithoutExtension(datPath);
                StringBuilder sb = new StringBuilder();
                sb.Append(ProcessHeader(protoName,groupName));
                sb.Append(ProcessVariables(workbook.GetSheetAt(0),protoName));
                sb.Append(ProcessMap(protoName));
                File.AppendAllText(Config.instance.protoDirectory+protoName+".proto",sb.ToString());
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                workbook.Close();
            }
        }
        private string ProcessHeader(string protoName,string groupName)
        {
            var info = Config.instance.GetGroupInfo(groupName);

            var header = "\nsyntax = \"proto3\";\noption csharp_namespace = \"{0}\";\n";
            return string.Format(header, info.namespacee);
        }

        private string ProcessVariables(ISheet sheet, string protoName)
        {
            StringBuilder sb = new StringBuilder("message " + protoName + " {\n");

            IRow row2 = sheet.GetRow(1);
            IRow row3 = sheet.GetRow(2);

            int count = 0;
            foreach (var item in row2)
            {
                string type = row2.GetCell(count).ToString();
                string name = row3.GetCell(count).ToString();
                count++;
                sb.Append(GetVariableString(count,type,name));
            }

            sb.Append("}");
            return sb.ToString();
        }

        private string GetVariableString(int index,string type, string name)
        {
            string str = "";
            if (type.Contains("[]"))
            {
                type = "repeated " + type.Split('[')[0];
            }
            str += " " + type + " " + name + " = " + index + ";";
            str += "\n";
            return str;
        }

        private string ProcessMap(string protoName)
        {
            string str = @"
message Excel_{0}
{{
    map<int32,{1}> {2} = 1;
}}";
            return string.Format(str, protoName, protoName, "Data");
        }
    }
}
