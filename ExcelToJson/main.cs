using System;
using System.Threading.Tasks;

namespace ExcelFormat
{
    class main
    {
        static void Main()
        {
            Formater formater = null;

            string[] args = Environment.GetCommandLineArgs();

            if (args[1].StartsWith("-"))
            {
                string[] option = args[1].Split('-');
                switch(option[1])
                {
                    case "json":
                        formater = new JsonFormater();
                        break;
                }
            }
            else
            {
                formater = new JsonFormater();
            }


            string[] filePaths = new string[args.Length -2];
            Array.Copy(args,2,filePaths,0,args.Length-2);

            Console.WriteLine("开始处理");
            try
            {
                formater.Format(filePaths);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadLine();
                return;
            }
            
            Console.WriteLine("共计" + filePaths.Length + "完成,文件路径为：");

            foreach (var item in filePaths)
            {
                Console.WriteLine(item);
            }
            Console.ReadLine();

        }
    }
}
