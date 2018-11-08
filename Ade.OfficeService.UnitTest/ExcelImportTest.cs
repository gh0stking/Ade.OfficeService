using Ade.OfficeService.Excel;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace Ade.OfficeService.UnitTest
{
    public class ExcelImportTest
    {
        [Fact]
        public void FilterTest()
        {
            string curDir = Environment.CurrentDirectory;

            string fileUrl = Path.Combine(curDir, "files", "CarImport.xlsx");

            CarImport carImportDTO = new CarImport();
            carImportDTO.SetDelegateDatabaseExist(DBExist);
            var rows = carImportDTO.Import(fileUrl);

            Assert.True(rows.Count > 0);

            var validRow = rows[5];
            CarImport dto = validRow.Convert<CarImport>();

            Assert.True(dto.CarCode == "¶õA17MG2" 
                && dto.Gender == GenderEnum.ÄÐ 
                && dto.Gender.GetHashCode() == 10
                && dto.Name == "¹¨Ó¢èº" 
                && dto.RegisterDate == new DateTime(2018, 1, 1) 
                && dto.Age == 18);
        }

        [Fact]
        public void ConvertTest()
        {

        }

        public bool DBExist(string tableName, string fieldName)
        {
            Assert.True(tableName == "cartable" && fieldName == "carcode");

            return true;
        }
    }
}
