using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class BaseFilterAttribute : Attribute
    {
        public string ErrorMsg { get; set; } = "校验失败";
    }
}
