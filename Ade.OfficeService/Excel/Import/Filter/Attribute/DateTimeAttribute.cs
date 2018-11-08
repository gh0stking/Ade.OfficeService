using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class DateTimeAttribute : BaseFilterAttribute
    {
        public DateTimeAttribute()
        {
            this.ErrorMsg = "不是有效日期";
        }
    }
}
