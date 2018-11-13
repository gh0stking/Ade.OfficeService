using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.ConsoleTest
{
    public class ImportCar 
    {
        [ColName("车牌号")]
        [Regex(RegexConstant.NOT_EMPTY_REGEX,ErrorMsg ="必填")]
        //[DatabaseExist("cartable","carcode")]
        [Regex(RegexConstant.CAR_CODE_REGEX)]
        //[Duplication]
        public string CarCode { get; set; }

        [ColName("手机号")]
        //[Regex(RegexEnum.国内手机号)]
        public string Mobile { get; set; }

        [ColName("身份证号")]
        [Regex(RegexConstant.IDENTITY_NUMBER_REGEX)]
        public string IdentityNumber { get; set; }

        [ColName("姓名")]
        [MaxLength(10)]
        public string Name { get; set; }

        [ColName("性别")]
        [Regex(RegexConstant.GENDER_REGEX)]
        public GenderEnum Gender { get; set; }

        [ColName("注册日期")]
        [DateTime]
        public DateTime RegisterDate { get; set; }

        [ColName("年龄")]
        [Range(0, 150)]
        public int Age { get; set; }
    }

    public enum GenderEnum
    {
        默认 = 0,
        男 = 10,
        女 = 20
    }
}
