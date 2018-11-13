using System;

namespace Ade.OfficeService
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColNameAttribute : Attribute
    {
        public ColNameAttribute(string colName)
        {
            this.ColName = colName;
        }

        public string ColName { get; set; }
    }
}
