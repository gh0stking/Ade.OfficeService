using System;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelExportAttribute : Attribute
    {
        public ExcelExportAttribute(string colName)
        {
            this.ColName = colName;
        }

        public string ColName { get; set; }
    }
}
