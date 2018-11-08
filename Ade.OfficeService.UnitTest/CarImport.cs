using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.UnitTest
{
    public class CarImport : BaseImport
    {
        [DatabaseExist("cartable","carcode")]
        [ExcelImport("车牌号")]
        [Regex(true, PreDefinedRegexEnum.车牌号)]
        public string CarCode { get; set; }

        [ExcelImport("姓名")]
        [MaxLength(10)]
        public string Name { get; set; }

        [ExcelImport("性别")]
        [Regex(true, PreDefinedRegexEnum.性别)]
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
