using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    //模板数据重复数据校验
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DuplicationAttribute : BaseFilterAttribute
    {
        public DuplicationAttribute()
        {
            this.ErrorMsg = "重复";
        }
    }
}
