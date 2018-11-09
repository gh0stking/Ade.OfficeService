using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Ade.OfficeService.Excel
{
    [FilterBind(typeof(RegexAttribute))]
    public class RegexFilter : IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context)
        {
            excelDataRows.Where(r => r.IsValid).ToList().ForEach(r => r.DataCols.ForEach(c => {
                var attrs = c.GetFilterAttrs<RegexAttribute>(context.TypeFilterInfo);

                if (attrs != null && attrs.Count > 0)
                {
                    attrs.ForEach(a =>
                    {
                        if (r.IsValid)
                        {
                            r.SetState(Regex.IsMatch(c.ColValue, a.Regex), c, a.ErrorMsg);
                        }
                    });
                }
            }));

            return excelDataRows;
        }
    }
}
