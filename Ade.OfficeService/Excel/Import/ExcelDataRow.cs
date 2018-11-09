using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class ExcelDataRow
    {
        public int RowIndex { get; set; }
        public List<ExcelDataCol> DataCols { get; set; } = new List<ExcelDataCol>();
        public bool IsValid { get; set; }
        public string ErrorMsg { get; set; }
    }
}
