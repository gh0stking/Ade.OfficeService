using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class TypeFilterInfoFactory
    {
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));
        public static TypeFilterInfo CreateInstance(Type importType, ExcelHeaderRow excelHeaderRow)
        {
            if (importType == null)
            {
                throw new ArgumentNullException("importDTOType");
            }

            if (excelHeaderRow == null)
            {
                throw new ArgumentNullException("excelHeaderRow");
            }

            var key = importType;
            if (Table[key] != null)
            {
                return (TypeFilterInfo)Table[key];
            }

            TypeFilterInfo typeFilterInfo = new TypeFilterInfo() { PropertyFilterInfos = new List<PropertyFilterInfo>() { } };

            IEnumerable<PropertyInfo> props = importType.GetProperties().ToList().Where(p => p.IsDefined(typeof(ExcelImportAttribute)));
            props.ToList().ForEach(p =>
            {
                string colName = p.GetCustomAttribute<ExcelImportAttribute>().ColName;
                ExcelCol col = excelHeaderRow.Cells.SingleOrDefault(c => c.ColName == colName);
                if (col != null)
                {
                    typeFilterInfo.PropertyFilterInfos.Add(
                        new PropertyFilterInfo
                        {
                            ColIndex = col.ColIndex,
                            FilterAttrs = p.GetCustomAttributes<BaseFilterAttribute>()?.ToList()
                        });
                }
            });

            Table[key] = typeFilterInfo;

            return typeFilterInfo;
        }
    }
}
