using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace Ade.OfficeService.UnitTest
{
    public class ExcelImportTest
    {
        [Fact]
        public void ImportTest()
        {
            string curDir = Environment.CurrentDirectory;

            string fileUrl = Path.Combine(curDir, "files", "CarImport.xlsx");

            ImportCar carImportDTO = new ImportCar();
            carImportDTO.SetDelegateDatabaseExist(DBExist);
            var rows = carImportDTO.Import(fileUrl);

            Assert.True(rows.Count > 0);
           
            var row0 = rows[0];
            Assert.True(!row0.IsValid && row0.ErrorMsg.Contains("车牌号非法"));

            var row1 = rows[1];
            Assert.True(!row1.IsValid && row1.ErrorMsg.Contains("姓名超长"));

            var row2 = rows[2];
            Assert.True(!row2.IsValid && row2.ErrorMsg.Contains("性别非法"));

            var row3 = rows[3];
            Assert.True(!row3.IsValid && row3.ErrorMsg.Contains("注册日期非法"));

            var row4 = rows[4];
            Assert.True(!row4.IsValid && row4.ErrorMsg.Contains("年龄超限，仅允许为0-150"));

            var row5 = rows[5];
            Assert.True(!row5.IsValid && row5.ErrorMsg.Contains("车牌号重复"));

            var row6 = rows[6];
            //Assert.True(!row6.IsValid && row6.ErrorMsgs.Any(e => e.Contains("已存在")));

            var row7 = rows[7];
            Assert.True(!row7.IsValid && row7.ErrorMsg.Contains("车牌号必填"));

            var row8 = rows[8];
            Assert.True(!row8.IsValid && row8.ErrorMsg.Contains("手机号非法"));

            var row9 = rows[9];
            Assert.True(!row9.IsValid && row9.ErrorMsg.Contains("身份证号非法"));

            var validRow = rows[10];
            ImportCar dto = validRow.Convert<ImportCar>();

            Assert.True(dto.CarCode == "鄂A57MG2"
                && dto.Gender == GenderEnum.男 
                && dto.Gender.GetHashCode() == 10
                && dto.Name == "龚英韬" 
                && dto.RegisterDate == new DateTime(2018, 1, 1) 
                && dto.Age == 18);
        }

        public bool DBExist(string tableName, string fieldName)
        {
            Assert.True(tableName == "cartable" && fieldName == "carcode");

            return true;
        }
    }
}
