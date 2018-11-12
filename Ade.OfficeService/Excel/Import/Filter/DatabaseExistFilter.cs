using System;
using System.Collections.Generic;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    [FilterBind(typeof(DatabaseExistAttribute))]
    public class DatabaseExistFilter : IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context)
        {
            if (context.DelegateNotExistInDatabase == null)
            {
                throw new ArgumentNullException("please set delegate of database filter first");
            }

            excelDataRows.Where(r => r.IsValid).ToList().ForEach(r => r.DataCols.ForEach(c =>
            {
                var attr = c.GetFilterAttr<DatabaseExistAttribute>(context.TypeFilterInfo);
                if (attr != null)
                {
                    r.SetState(context.DelegateNotExistInDatabase(
                        new DatabaseFilterContext()
                    {
                        TableName = attr.TableName,
                        FieldName = attr.FieldName,
                        Value = c.ColValue
                    }), c, attr.ErrorMsg);
                }
            }));

            return excelDataRows;
        }
    }
}
