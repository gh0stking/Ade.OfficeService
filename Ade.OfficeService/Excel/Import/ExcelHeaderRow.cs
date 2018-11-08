using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class ExcelHeaderRow
    {
        public List<ExcelCol> Cells { get; set; } = new List<ExcelCol>();
    }
}
