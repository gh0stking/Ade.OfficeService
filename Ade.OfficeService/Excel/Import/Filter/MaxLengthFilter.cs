using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    [FilterBind(typeof(MaxLengthAttribute))]
    public class MaxLengthFilter : IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context)
        {
            excelDataRows.Where(r => r.IsValid).ToList().ForEach(r => r.DataCols.ForEach(c => {
                var attr = c.GetFilterAttr<MaxLengthAttribute>(context.TypeFilterInfo);
                if (attr != null)
                {
                    r.SetState(c.ColValue.Length<=attr.MaxLength, c, attr.ErrorMsg);
                }
            }));

            return excelDataRows;
        }
    }
}
