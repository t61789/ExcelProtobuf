using Google.Protobuf;
using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelProtobuf
{
    public class DataConverter
    {
        public void Process(string groupName)
        {
            Program.Log("***开始转换数据***");

            var info = Config.instance.GetGroupInfo(groupName);
            Assembly assembly = Assembly.LoadFrom(info.dllDir+Path.DirectorySeparatorChar+info.namespacee+Config.instance.GetDllExtension(groupName));

            foreach (var (datPath, excelName) in Config.instance.GetMappings(groupName))
            {
                try
                {
                    ProcessData(assembly, datPath, excelName, info.namespacee);
                }
                catch (Exception e)
                {
                    Config.instance.RecordLog(e);
                    Program.Log("数据转换失败 {0}", datPath);
                    continue;
                }
                Program.Log("数据转换成功 {0}", datPath);
            }
        }

        private void ProcessData(Assembly assembly, string dataPath, string excelPath,string namespacee)
        {
            string dataName = Path.GetFileNameWithoutExtension(dataPath);

            Type serializerType = assembly.GetType(namespacee+".Excel_"+dataName);
            Type rowType = assembly.GetType(namespacee +"."+dataName);

            XSSFWorkbook workbook;
            using(FileStream fs = new FileStream(Config.instance.excelDirectory+excelPath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fs);
            }
            ISheet sheet =workbook.GetSheetAt(0);

            Dictionary<string, int> fieldMap = new Dictionary<string, int>();
            List<ICell> fieldRow = sheet.GetRow(2).Cells;
            for (int i = 0; i < fieldRow.Count; i++)
                fieldMap.Add(Program.FirstCharUpper(fieldRow[i].ToString()), i);

            object newTable = serializerType.GetConstructor(new Type[0]).Invoke(null);

            Type dic = serializerType.GetProperty("Data").PropertyType;
            object diction = serializerType.GetProperty("Data").GetValue(newTable);
            MethodInfo addMethod = dic.GetMethod("Add",new Type[] { typeof(int), rowType });

            int count = 0;
            foreach (IRow item in sheet)
            {
                if (count < 3)
                {
                    count++;
                    continue;
                }
                object newRow = rowType.GetConstructor(new Type[0]).Invoke(null);
                foreach (var i in rowType.GetProperties(BindingFlags.Public|BindingFlags.Instance))
                {
                    object value = Convert.ChangeType(item.GetCell(fieldMap[i.Name])?.ToString(), i.PropertyType);
                    if (value == null)
                        value = Default(i.PropertyType);

                    i.SetValue(newRow, value);
                }

                addMethod.Invoke(diction,new object[] {count-3,newRow });
                count++;
            }

            workbook.Close();

            using(FileStream fs = new FileStream(dataPath,FileMode.Create, FileAccess.Write))
            {
                byte[] result = new byte[(int)serializerType.GetMethod("CalculateSize").Invoke(newTable,null)];
                CodedOutputStream o = new CodedOutputStream(result);
                serializerType.GetMethod("WriteTo").Invoke(newTable,new object[] { o });
                fs.Write(result,0,result.Length);
            }
        }

        public object Default(Type type)
        {
            if (type == typeof(double))
                return 0.0;
            else if (type == typeof(float))
                return 0f;
            else if (type == typeof(int))
                return 0;
            else if (type == typeof(long))
                return 0;
            else if (type == typeof(uint))
                return 0;
            else if (type == typeof(ulong))
                return 0;
            else if (type == typeof(bool))
                return false;
            else if (type == typeof(string))
                return "";
            else if (type == typeof(byte[]))
                return new byte[0];
            return null;
        }
    }
}