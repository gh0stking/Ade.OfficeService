using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Ade.OfficeService.Excel
{
    [FilterBindOption(typeof(RangeAttribute))]
    public class RangeFilter : IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context)
        {
            excelDataRows.Where(r => r.IsValid).ToList().ForEach(r => r.DataCols.ForEach(c =>
            {
                var attr = c.GetFilterAttr<RangeAttribute>(context.TypeAttrsFlyWeight);
                if (attr != null)
                {
                    r.SetState(c.IsInRange(attr.Max, attr.Min), c, attr.ErrorMsg);
                }
            }));

            return excelDataRows;
        }
    }
}
