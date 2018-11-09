using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    [FilterBind(typeof(DateTimeAttribute))]
    public class DateTimeFilter : IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context)
        {
            excelDataRows.Where(r => r.IsValid).ToList().ForEach(r => r.DataCols.ForEach(c => {
                var attr = c.GetFilterAttr<DateTimeAttribute>(context.TypeFilterInfo);
                if (attr != null)
                {
                    r.SetState(c.IsDateTime(), c, attr.ErrorMsg);
                }
            }));

            return excelDataRows;
        }
    }
}
