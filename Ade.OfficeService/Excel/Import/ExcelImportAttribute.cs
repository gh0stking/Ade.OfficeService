using System;

namespace Ade.OfficeService
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelImportAttribute : Attribute
    {
        public ExcelImportAttribute(string colName)
        {
            this.ColName = colName;
        }

        public string ColName { get; set; }
    }
}
