using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Ade.OfficeService.ConsoleTest
{
    public class ExcelImportTest
    {
        public static void ImportTest()
        {
            string curDir = Environment.CurrentDirectory;

            string fileUrl = Path.Combine(curDir, "files", "CarImport.xlsx");

            var rows = ExcelImportService.Import<ImportCar>(fileUrl, DBExist);

            List<ImportCar> list = new List<ImportCar>();
            foreach (var item in rows.Where(e => e.IsValid))
            {
                //����ת�� - 5000�� 6��
                //list.Add(item.Convert<ImportCar>());

                //Expression + ����ת�� - 5000��3.5��
                list.Add(ExpressionMapper.FastConvert<ImportCar>(item));
            }
        }

        public static bool DBExist(DatabaseFilterContext context)
        {
            return true;
        }
    }
}
