using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProtobuf
{
    public class Program
    {
        private static Process p = new Process();

        [STAThread]
        public static void Main(string[] args)
        {
            new Program().Start();
        }

        public void Start()
        {
            while (true)
            {
                int command = Command("开始使用", "打开表文件", "添加新的映射", "更新所有数据","强制更新所有数据", "删除闲置映射和缓存", "打开配置文件","退出");
                switch (command)
                {
                    case 1:
                        OpenFile();
                        break;
                    case 2:
                        AddNewFile();
                        break;
                    case 3:
                        UpdateAllFile(false);
                        break;
                    case 4:
                        UpdateAllFile(true);
                        break;
                    case 5:
                        DeleteEmptyMapping();
                        break;
                    case 6:
                        OpenConfigFile();
                        break;
                    case 7:
                        Environment.Exit(0);
                        break;
                    default:
                        break;
                }
            }
        }

        public void AddNewFile()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "bytes二进制文件|*.bytes|dat数据库|*.dat|所有文件|*.*",
                Title = "保存表文件",
            };
            DialogResult d = saveFileDialog.ShowDialog();
            if (d == DialogResult.OK)
            {
                string datPath = saveFileDialog.FileName;
                Console.Write("组名:>");
                string group = Console.ReadLine();

                if(Config.instance.AddNewMappingAndExcel(datPath, group, true))
                    Log("映射创建成功 {0}",datPath);
            }
        }

        public void OpenFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "bytes二进制文件|*.bytes|dat数据库|*.dat|所有文件|*.*",
                Title = "保存表文件"
            };

            DialogResult d = openFileDialog.ShowDialog();
            if (d == DialogResult.OK)
            {
                string excelPath = Config.instance.GetMapping(openFileDialog.FileName);
                if (excelPath == null)
                {
                    Log("到 {0} 的映射不存在",openFileDialog.FileName);
                    return;
                }
                else
                {
                    excelPath = Config.instance.excelDirectory + excelPath;
                    OpenFile(excelPath);
                }
            }
        }

        public void DeleteEmptyMapping()
        {
            string[] del = Config.instance.GetEmptyMapping();
            Log("开始扫描闲置映射……");
            if (del.Length == 0)
            {
                Log("无闲置映射");
            }
            else
            {
                foreach (var item in del)
                    Console.WriteLine(item);
                int command = Command("是否删除以上闲置映射及其Excel缓存?","是","否");
                if (command == 1)
                {
                    Config.instance.DeleteEmptyMapping();
                }
            }

            Log("开始扫描闲置Excel缓存……");
            del = Config.instance.GetInvalidExcel();
            if (del.Length == 0)
            {
                Log("无闲置Excel缓存");
            }
            else
            {
                foreach (var item in del)
                    Console.WriteLine(item);
                int command = Command("是否删除以上Excel缓存?", "是", "否");
                if (command == 1)
                {
                    Config.instance.DeleteInvalidExcel(del);
                }
            }
        }

        public void OpenConfigFile()
        {
            OpenFile(Config.instance.configPath);
        }

        public void UpdateAllFile(bool force)
        {
            Config.instance.LoadConfig();
            string[] groups = Config.instance.GetGroups();
            ProtoHandler protoHandler = new ProtoHandler();
            Compiler compiler = new Compiler();
            DataConverter dataHandler = new DataConverter();

            foreach (var item in groups)
            {
                Log("[组 {0} 开始转换]", item);
                try
                {
                    protoHandler.Process(item, force);
                    compiler.Process(item);
                    dataHandler.Process(item);
                }
                catch (Exception e)
                {
                    Config.instance.RecordLog(e);
                    Log("[组 {0} 转换失败]", item);
                    continue;
                }
                Log("[组 {0} 转换完毕]", item);
            }
        }

        public static int Command(string discripe, params string[] selection)
        {
            while (true)
            {
                Console.WriteLine("\n[" + discripe+"]");
                for (int i = 1; i <= selection.Length; i++)
                    Console.WriteLine(i.ToString() + '.' + selection[i - 1]);
                Console.Write("command:>");
                if (int.TryParse(Console.ReadLine(), out int result))
                {
                    if (result >= 1 && result <= selection.Length)
                        return result;
                }
                Console.WriteLine();
            }
        }

        public static void OpenFile(string path)
        {
            p.StartInfo = new ProcessStartInfo()
            {
                FileName = path
            };
            p.Start();
            p.Close();
        }

        public static void Exec(string file,string arg)
        {
            p.StartInfo = new ProcessStartInfo()
            {
                FileName = file,
                Arguments = arg,
                UseShellExecute = false,
                RedirectStandardOutput = true
            };
            p.Start();
            p.WaitForExit();
            p.StandardOutput.ReadToEnd();
            p.Close();
        }

        public static void Log(string message,params object[] content)
        {
            Console.WriteLine(DateTime.Now+" :>"+message,content);
        }

        public static string FirstCharUpper(string str)
        {
            if (str == "" || str == null)
                return str;

            return char.ToUpper(str[0]) + str.Substring(1);
        }
    }
}
