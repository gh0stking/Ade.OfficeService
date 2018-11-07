using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColNameAttribute : Attribute
    {
        public ColNameAttribute(string colName)
        {
            ColName = colName;
        }
        public string ColName { get; set; }
    }
}
