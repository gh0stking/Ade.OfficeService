using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class MaxLengthAttribute : BaseFilterAttribute
    {
        public MaxLengthAttribute(int maxLength)
        {
            this.MaxLength = maxLength;
            this.ErrorMsg = "超长";
        }

        public int MaxLength { get; set; }
    }
}
