using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Ade.OfficeService.Excel
{
    [FilterBindOption(typeof(RegexAttribute))]
    public class RegexFilter : IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context)
        {
            excelDataRows.Where(r => r.IsValid).ToList().ForEach(r => r.DataCols.ForEach(c => {
                var attr = c.GetFilterAttr<RegexAttribute>(context.TypeAttrsFlyWeight);
                if (attr != null)
                {
                    r.SetState(Regex.IsMatch(c.ColValue, attr.Regex), c, attr.ErrorMsg);
                }
            }));

            return excelDataRows;
        }
    }
}
