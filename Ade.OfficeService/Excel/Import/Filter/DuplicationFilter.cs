using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    [FilterBind(typeof(DuplicationAttribute))]
    public class ValidateDuplication : IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context)
        {
            List<KeyValuePair<int, string>> kvps = new List<KeyValuePair<int, string>>();
            excelDataRows.ToList().ForEach(r => r.DataCols.ForEach(c =>
            {
                KeyValuePair<int, string> kvp = new KeyValuePair<int, string>(c.ColIndex, c.ColValue);
                if (r.IsValid)
                {
                    var attr = c.GetFilterAttr<DuplicationAttribute>(context.TypeFilterInfo);
                    if (attr != null)
                    {
                        r.SetState(!kvps.Contains(kvp), c, attr.ErrorMsg);
                    }
                }

                kvps.Add(kvp);
            }));

            return excelDataRows;
        }
    }
}
