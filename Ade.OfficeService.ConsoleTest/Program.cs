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
            var rows = ExcelImportService.Import<ImportCar>(fileUrl, DBExist);
            sw.Stop();

            Console.WriteLine($"Exel读取以及校验耗时：{sw.ElapsedMilliseconds}");

            List<ImportCar> list = new List<ImportCar>();

            sw.Restart();
            foreach (var item in rows.Where(e => e.IsValid))
            {
                //Expression + 缓存转换 - 5000条3.5秒
                list.Add(HardCode(item));
            }
            sw.Stop();
            Console.WriteLine($"硬编码转换耗时：{sw.ElapsedMilliseconds}");

            sw.Restart();
            foreach (var item in rows.Where(e=>e.IsValid))
            {
                //反射转换 - 5000条 6秒
               item.Convert<ImportCar>();
            }
            sw.Stop();
            Console.WriteLine($"反射+委托转换耗时：{sw.ElapsedMilliseconds}");

            sw.Restart();
            foreach (var item in rows.Where(e => e.IsValid))
            {
                //反射转换 - 5000条 6秒
                item.ConvertDirect<ImportCar>();
            }
            sw.Stop();
            Console.WriteLine($"直接反射转换耗时：{sw.ElapsedMilliseconds}");


            sw.Restart();
            foreach (var item in rows.Where(e=>e.IsValid))
            {
                //Expression + 缓存转换 - 5000条3.5秒
                list.Add(ExpressionMapper.Trans<ImportCar>(item));
            }
            sw.Stop();
            Console.WriteLine($"Expression转换耗时(无缓存)：{sw.ElapsedMilliseconds}");

            Console.WriteLine("---------------------------------------------------------");
        }

        public static ImportCar HardCode(ExcelDataRow row)
        {
            return new ImportCar()
            {
                Age = int.Parse(row.DataCols.SingleOrDefault(c => c.PropertyName == "Age").ColValue),
                CarCode = row.DataCols.SingleOrDefault(c => c.PropertyName == "CarCode").ColValue,
                Gender = (GenderEnum)Enum.Parse(typeof(GenderEnum), row.DataCols.SingleOrDefault(c => c.PropertyName == "Gender").ColValue),
                IdentityNumber = row.DataCols.SingleOrDefault(c => c.PropertyName == "IdentityNumber").ColValue,
                Mobile = row.DataCols.SingleOrDefault(c => c.PropertyName == "Mobile").ColValue,
                Name = row.DataCols.SingleOrDefault(c => c.PropertyName == "Name").ColValue,
                RegisterDate = DateTime.Parse(row.DataCols.SingleOrDefault(c => c.PropertyName == "RegisterDate").ColValue)
            };
        }

        public static bool DBExist(DatabaseFilterContext context)
        {
            return true;
        }
    }
}
