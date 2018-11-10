using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class ExcelDataCol : ExcelCol
    {
        internal string PropertyName { get; set; }
        public int RowIndex { get; set; }
        public string ColValue { get; set; }
    }
}
