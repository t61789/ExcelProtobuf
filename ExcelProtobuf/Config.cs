using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelProtobuf
{
    public class Config
    {
        public static Config instance;

        public string configPath;
        public string excelDirectory;
        public string protoDirectory;
        public string codeDirectory;
        public string protoDllPath;
        public string protocPath;
        public string logPath;

        public string debugPath;

        public XDocument configDoc;

        static Config()
        {
            instance = new Config
            {
                configPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"config.xml",
                excelDirectory = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"excel" + Path.DirectorySeparatorChar,
                protoDirectory = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"protos" + Path.DirectorySeparatorChar,
                codeDirectory = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"codes" + Path.DirectorySeparatorChar,
                protoDllPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"Google.Protobuf.dll",
                protocPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"protoc-3.8.0-win64" + Path.DirectorySeparatorChar + @"bin" + Path.DirectorySeparatorChar + @"protoc.exe",
                logPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"log.log",
                debugPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"debug.log",
            };
            instance.LoadConfig();
        }

        public void LoadConfig()
        {
            configDoc = XDocument.Load(configPath);
        }

        public void SaveConfig()
        {
            configDoc.Save(configPath);
        }

        public (string datPath, string excelName)[] GetMappings(string groupName)
        {
            XElement e = configDoc.Root.Element("mapping").Elements().Where(x => x.Attribute("name").Value == groupName).FirstOrDefault();
            if (e == null)
            {
                Console.WriteLine("组 {0} 不存在", groupName);
                throw new Exception();
            }

            List<(string datPath, string excelName)> result = new List<(string, string)>();
            foreach (var item in e.Elements())
                result.Add((item.Element("dat").Value, item.Element("excel").Value));
            return result.ToArray();
        }

        public string GetMapping(string dat)
        {
            foreach (var item in configDoc.Root.Element("mapping").Elements())
            {
                string e = item.Elements().Where(x => x.Element("dat").Value == dat).FirstOrDefault()?.Element("excel").Value;
                if (e != null) return e;
            }
            return null;
        }

        public bool ExcelExists(string excelName)
        {
            foreach (var item in configDoc.Root.Element("mapping").Elements())
            {
                foreach (var i in item.Elements())
                {
                    if (i.Element("excel").Value == excelName)
                        return true;
                }
            }

            return false;
        }

        public string GetComplierPath()
        {
            return configDoc.Root.Element("complier").Value;
        }

        public string GetHash(string fileName)
        {
            string sha256;
            SHA256Managed s = new SHA256Managed();
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                sha256 = BitConverter.ToString(s.ComputeHash(fs));
            }
            return sha256;
        }

        public bool AddNewMappingAndExcel(string dataPath, string groupName,bool cover)
        {
            if (groupName == "" || groupName == null) groupName = "default";

            XElement group = configDoc.Root.Element("mapping").Elements().Where(x => x.Attribute("name").Value == groupName).FirstOrDefault();
            if (group == null)
            {
                Program.Log("组 {0} 不存在",groupName);
                return false;
            }

            XElement tempE = group.Elements().Where(x => x.Element("dat").Value == dataPath).FirstOrDefault();
            if (tempE != null)
            {
                if (cover)
                {
                    File.Delete(dataPath);
                    File.Delete(excelDirectory+ tempE.Element("excel").Value);
                    tempE.Remove();
                }else
                {
                    Program.Log("组 {0} 已存在 {1} 的映射",groupName,dataPath);
                    return false;
                }
            }

            string excelName = Path.GetFileName(dataPath);
            int lastIndex = excelName.LastIndexOf('.');
            if (lastIndex == -1)
                excelName += ".xlsx";
            else
                excelName = excelName.Substring(0,lastIndex)+ ".xlsx";

            string curdir = excelName;
            while (ExcelExists(excelName))
            {
                curdir = Path.GetDirectoryName(curdir);
                if (curdir == "")
                {
                    Program.Log("Excel缓存重名");
                    return false;
                }
                string temp = Path.GetFileName(curdir);
                if (temp == "")
                    temp = curdir[0].ToString();

                excelName = temp + '_' + excelName;
            }

            File.Create(dataPath).Close();

            XSSFWorkbook workbook = new XSSFWorkbook();
            workbook.CreateSheet();
            using (FileStream fs = new FileStream(excelDirectory + excelName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            workbook.Close();

            XElement datE = new XElement("dat") { Value = dataPath };
            XElement excelE = new XElement("excel") { Value = excelName };
            XElement map = new XElement("map");
            map.Add(datE);
            map.Add(excelE);
            group.Add(map);

            SaveConfig();

            return true;
        }

        public (string dllDir, string namespacee) GetGroupInfo(string groupName)
        {
            XElement e = configDoc.Root.Element("mapping").Elements().Where(x => x.Attribute("name").Value == groupName).FirstOrDefault();
            if (e == null)
                return (null, null);
            else
                return (e.Attribute("dllDir").Value, e.Attribute("namespace").Value);
        }

        public string[] GetGroups()
        {
            List<string> result = new List<string>();
            foreach (var item in configDoc.Root.Element("mapping").Elements())
                result.Add(item.Attribute("name").Value);

            return result.ToArray();
        }

        public string[] GetEmptyMapping()
        {
            List<string> result = new List<string>();
            foreach (var item in configDoc.Root.Element("mapping").Elements())
            {
                foreach (var i in item.Elements())
                {
                    string temp = i.Element("dat").Value;
                    if (!File.Exists(temp))
                    {
                        result.Add(temp);
                    }
                }
            }
            return result.ToArray();
        }

        public void DeleteEmptyMapping()
        {
            List<XElement> result = new List<XElement>();
            foreach (var item in configDoc.Root.Element("mapping").Elements())
            {
                foreach (var i in item.Elements())
                {
                    string temp = i.Element("dat").Value;
                    if (!File.Exists(temp))
                    {
                        result.Add(i);
                    }
                }
            }
            foreach (var item in result)
            {
                try
                {
                    File.Delete(excelDirectory + item.Element("excel").Value);
                    Program.Log("删除成功 {0}", item.Element("dat").Value);
                    item.Remove();
                }
                catch (Exception)
                {
                    Program.Log("删除失败 {0}", item.Element("dat").Value);
                    continue;
                }
            }

            SaveConfig();
        }

        public bool GroupHashCheck(string groupName)
        {
            string temp = "";
            XElement tempe = configDoc.Root.Element("mapping").Elements().Where(x => x.Attribute("name").Value == groupName).FirstOrDefault();
            foreach (var item in tempe.Elements())
            {
                SHA256Managed s = new SHA256Managed();
                using (FileStream fs = new FileStream(excelDirectory + item.Element("excel").Value, FileMode.Open, FileAccess.Read))
                {
                    string fileHash = BitConverter.ToString(s.ComputeHash(fs)).Replace("-", "");
                    temp = BitConverter.ToString(s.ComputeHash(Encoding.ASCII.GetBytes(temp + fileHash))).Replace("-", "");
                }
            }

            bool result = tempe.Attribute("hash").Value.Equals(temp);
            if(!result)
            {
                tempe.Attribute("hash").Value = temp;
                SaveConfig();
            }
            return result;
        }

        public string[] GetInvalidExcel()
        {
            List<string> result = new List<string>();
            foreach (var item in Directory.GetFiles(excelDirectory,"*.xlsx"))
            {
                bool flag = true;
                foreach (var i in configDoc.Root.Element("mapping").Elements())
                {
                    XElement e = i.Elements().Where(x => x.Element("excel").Value == Path.GetFileName(item)).FirstOrDefault();
                    if (e != null)
                    {
                        flag = false;
                        break;
                    }
                }
                if (flag)
                    result.Add(item);
            }
            return result.ToArray();
        }

        public void DeleteInvalidExcel(string[] excelPaths)
        {
            foreach (var item in excelPaths)
            {
                try
                {
                    File.Delete(item);
                    Program.Log("删除成功 {0}",item);
                }
                catch (Exception)
                {
                    Program.Log("删除失败 {0}", item);
                }
            }
        }

        public string GetDllExtension(string groupName)
        {
            string extension = configDoc.Root.Element("mapping").Elements().Where(x => x.Attribute("name").Value == groupName).FirstOrDefault()?.Attribute("dll-extension").Value;
            if (extension == "") extension = ".dll";
            return extension;
        }

        public void RecordLog(Exception e)
        {
            if (!File.Exists(logPath))
                File.Create(logPath).Close();
            string s = string.Format("[{0}] :>{1}\n", DateTime.Now, e);
            File.AppendAllText(logPath,s);
        }

        public void Debugg(string str)
        {
            if (!File.Exists(debugPath))
                File.Create(debugPath).Close();
            string s = string.Format("[{0}] :>{1}\n", DateTime.Now, str);
            File.AppendAllText(debugPath, s);
        }
    }
}
