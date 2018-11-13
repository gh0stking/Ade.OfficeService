using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.UnitTest
{
    [WrapText]
    [Header(Color =ColorEnum.RED,FontName ="微软雅黑",FontSize =12,IsBold =true)]
    public class ExportCar
    {
        [MergeCols]
        [ColName("车牌号")]
        public string CarCode { get; set; }

        [ColName("姓名")]
        public string Name { get; set; }

        [ColName("性别")]
        public GenderEnum Gender { get; set; }

        [ColName("注册日期")]
        public DateTime RegisterDate { get; set; }

        [ColName("年龄")]
        public int Age { get; set; }
    }
}
