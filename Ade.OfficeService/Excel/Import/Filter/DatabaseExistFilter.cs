using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    [FilterBindOption(typeof(DatabaseExistAttribute))]
    public class DatabaseExistFilter : IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context)
        {
            if (context.DelegateDatabaseFilter == null)
            {
                throw new ArgumentNullException("please set delegate of database filter first");
            }

            excelDataRows.Where(r => r.IsValid).ToList().ForEach(r => r.DataCols.ForEach(c =>
            {
                var attr = c.GetFilterAttr<DatabaseExistAttribute>(context.TypeAttrsFlyWeight);
                if (attr != null)
                {
                    r.SetState(context.DelegateDatabaseFilter(attr.TableName, attr.FieldName), c, attr.ErrorMsg);
                }
            }));

            return excelDataRows;
        }
    }
}
