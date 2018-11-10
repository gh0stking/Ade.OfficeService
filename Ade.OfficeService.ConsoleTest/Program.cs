using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Ade.OfficeService.ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Test();
                Console.ReadKey();
            }
        }

        private static void Test()
        {

            string curDir = Environment.CurrentDirectory;

            string fileUrl = Path.Combine(curDir, "files", "CarImport.xlsx");

            Stopwatch sw = new Stopwatch();

            sw.Start();
            var rows = ExcelImportService<ImportCar>.Import(fileUrl, DBExist);
            sw.Stop();

            Console.WriteLine($"校验耗时：{sw.ElapsedMilliseconds}");

            sw.Restart();
            List<ImportCar> list = new List<ImportCar>();
            foreach (var item in rows)
            {
                //反射转换 - 5000条 6秒
               // item.Convert<ImportCar>();
            }
            sw.Stop();
            Console.WriteLine($"反射转换耗时：{sw.ElapsedMilliseconds}");

            sw.Restart();
            foreach (var item in rows)
            {
                //Expression + 缓存转换 - 5000条3.5秒
                list.Add(ExpressionMapper.Trans<ImportCar>(item));
            }
            sw.Stop();
            Console.WriteLine($"Expression转换耗时：{sw.ElapsedMilliseconds}");
        }

        public static bool DBExist(DatabaseFilterContext context)
        {
            return true;
        }
    }
}
