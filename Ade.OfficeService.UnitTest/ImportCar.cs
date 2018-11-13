using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.UnitTest
{
    public class ImportCar : IExcelImport
    {
        [ExcelImport("车牌号")]
        [Regex(RegexConstant.NOT_EMPTY_REGEX,ErrorMsg ="必填")]
        [DatabaseExist("cartable","carcode")]
        [Regex(RegexConstant.CAR_CODE_REGEX)]
        //[Duplication]
        public string CarCode { get; set; }

        [ExcelImport("手机号")]
        [Regex(RegexConstant.MOBILE_CHINA_REGEX)]
        public string Mobile { get; set; }

        [ExcelImport("身份证号")]
        [Regex(RegexConstant.IDENTITY_NUMBER_REGEX)]
        public string IdentityNumber { get; set; }

        [ExcelImport("姓名")]
        [MaxLength(10)]
        public string Name { get; set; }

        [ExcelImport("性别")]
        [Regex(RegexConstant.GENDER_REGEX)]
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
