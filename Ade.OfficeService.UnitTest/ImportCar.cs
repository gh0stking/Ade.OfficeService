using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.UnitTest
{
    public class ImportCar : BaseExcelImport
    {
        [ExcelImport("车牌号")]
        [Regex(RegexEnum.非空,ErrorMsg ="必填")]
        [DatabaseExist("cartable","carcode")]
        [Regex(RegexEnum.车牌号)]
        [Duplication]
        public string CarCode { get; set; }

        [ExcelImport("手机号")]
        [Regex(RegexEnum.国内手机号)]
        public string Mobile { get; set; }

        [ExcelImport("身份证号")]
        [Regex(RegexEnum.身份证号)]
        public string IdentityNumber { get; set; }

        [ExcelImport("姓名")]
        [MaxLength(10)]
        public string Name { get; set; }

        [ExcelImport("性别")]
        [Regex(RegexEnum.性别)]
        public GenderEnum Gender { get; set; }

        [ExcelImport("注册日期")]
        [DateTime]
        public DateTime RegisterDate { get; set; }

        [ExcelImport("年龄")]
        [Range(0, 150)]
        public int Age { get; set; }
    }

    public enum GenderEnum
    {
        男 = 10,
        女 = 20
    }
}
