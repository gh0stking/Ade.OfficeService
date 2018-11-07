using System;

namespace Ade.OfficeService
{
    public class ColNameAttribute : Attribute
    {
        public ColNameAttribute(string colName)
        {
            this.ColName = colName;
        }

        public string ColName { get; set; }
    }
}
