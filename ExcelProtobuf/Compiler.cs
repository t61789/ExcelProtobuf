using System.IO;

namespace ExcelProtobuf
{
    public class Compiler
    {
        public void Process(string groupName)
        {
            Program.Log("***开始生成编译dll***");

            string path = Config.instance.codeDirectory;
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
            Directory.CreateDirectory(path);

            int count = 0;
            foreach (var item in Directory.GetFiles(Config.instance.protoDirectory, "*.proto"))
            {
                Program.Exec(Config.instance.protocPath, string.Format(" --proto_path={0} --csharp_out={1} {2}", Config.instance.protoDirectory, Config.instance.codeDirectory, item));
                string fileName = Path.GetFileNameWithoutExtension(item);

                int temp = Directory.GetFiles(Config.instance.codeDirectory, "*.cs").Length;
                if (count+1 == temp)
                    Program.Log("转换为cs文件成功 {0}", fileName);
                else
                    Program.Log("转换为cs文件失败 {0}", fileName);
                count = temp;
            }

            Compile(groupName);
        }

        public void Compile(string groupName)
        {
            Program.Log("开始生成dll文件");

            var info = Config.instance.GetGroupInfo(groupName);
            var command = @"-target:library -out:{0} -reference:{1} -recurse:{2}\*.cs";
            var dllPath = info.dllDir+Path.DirectorySeparatorChar + info.namespacee+Config.instance.GetDllExtension(groupName);
            var csharpFolder = Config.instance.codeDirectory;
            Program.Exec(Config.instance.GetComplierPath(), string.Format(command, dllPath, Config.instance.protoDllPath, csharpFolder));

            if (!File.Exists(dllPath))
            {
                Program.Log("dll文件生成失败");
                throw new System.Exception();
            }

            Program.Log("dll文件生成成功 {0}", dllPath);
        }
    }
}
