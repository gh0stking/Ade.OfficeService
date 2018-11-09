using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DateTimeAttribute : BaseFilterAttribute
    {
        public DateTimeAttribute()
        {
            this.ErrorMsg = "非法";
        }
    }
}
