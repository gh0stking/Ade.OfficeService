using Ade.OfficeService.Excel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace Ade.OfficeService.UnitTest
{
    public class ExcelExportTest
    {
        [Fact]
        public void ExportTest()
        {
            List<ExportCar> list = new List<ExportCar>();

            for (int i = 0; i < 5; i++)
            {
                list.Add(new ExportCar()
                {
                    Name = "龚英韬",
                    Age = 22,
                    CarCode = "SSSS",
                    Gender = GenderEnum.男,
                    RegisterDate = DateTime.Now
                });
            }

            for (int i = 0; i < 600; i++)
            {
                list.Add(new ExportCar()
                {
                    Name = "龚英韬111",
                    Age = 22,
                    CarCode = "对对对对对对多多多多多多多多多多多多多多多多多多多多多多多多多多多多多多多多多多多多",
                    Gender = GenderEnum.男,
                    RegisterDate = DateTime.Now
                });
            }

            IWorkbook wk = ExcelExportService<ExportCar>.Export(list);
            Assert.True(wk != null);


            File.WriteAllBytes(@"c:\test\test.xls", wk.ToBytes());
        }
       
    }
}
