using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class RangeAttribute : BaseFilterAttribute
    {
        public int Min { get; set; }
        public int Max { get; set; }

        public RangeAttribute(int min, int max)
        {
            this.Min = min;
            this.Max = max;
            this.ErrorMsg = $"仅允许{min} - {max}的值";
        }
    }
}
