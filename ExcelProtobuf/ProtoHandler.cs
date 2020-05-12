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
        private const string TYPES = "double float int32 int64 uint32 uint64 sint32 sint64 fixed32 fixed64 sfixed32 sfixed64 bool string bytes";

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

            var header = "syntax = \"proto3\";\noption csharp_namespace = \"{0}\";\n";
            return string.Format(header, info.namespacee);
        }

        private string ProcessVariables(ISheet sheet, string protoName)
        {
            StringBuilder sb = new StringBuilder("message " + protoName + " {\n");

            IRow row2 = sheet.GetRow(1);

            int count = 0;
            foreach (var item in row2)
            {
                string type = row2.GetCell(count).ToString();
                count++;
                sb.Append(GetVariableString(count,type));
            }

            sb.Append("}");
            return sb.ToString();
        }

        private string GetVariableString(int index,string type)
        {
            string s = "{0} {1} D{2} = {2};\n";

            int result = TypeCheck(type);
            if (result == 0) throw new InvalidCastException("proto类型不正确: "+ type);
            else if(result == 1)
                s = string.Format(s, null, type, index);
            else if(result == 2)
                s = string.Format(s, "repeated", type.Split('[')[0], index);

            return s;
        }

        private string ProcessMap(string protoName)
        {
            string str = @"
message Excel_{0}
{{
    repeated string Fields = 1;
    repeated {0} Data = 2;
}}";
            return string.Format(str, protoName);
        }

        private int TypeCheck(string type)
        {
            if (TYPES.Contains(type)) return 1;
            string[] spli = type.Split('[');
            if (spli.Length != 2) return 0;
            if (spli[1] != "]") return 0;
            if (TYPES.Contains(spli[0])) return 2;
            return 0;
        }
    }
}
