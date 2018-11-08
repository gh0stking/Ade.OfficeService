using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class TypeFilterAttrsFlyWeight
    {
        private TypeFilterAttrsFlyWeight()
        {
        }
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));
        public static TypeFilterAttrsFlyWeight CreateInstance(Type importDTOType,ExcelHeaderRow excelHeaderRow)
        {
            if (importDTOType == null)
            {
                throw new ArgumentNullException("importDTOType");
            }

            if (excelHeaderRow == null)
            {
                throw new ArgumentNullException("excelHeaderRow");
            }

            var key = new TypeFilterAttrsFlyWeightKey { ImportDTOType = importDTOType, ExcelHeaderRow = excelHeaderRow };
            if (Table[key] != null)
            {
                return (TypeFilterAttrsFlyWeight)Table[key];
            }

            TypeFilterAttrsFlyWeight typeAttrsFlyWeight = new TypeFilterAttrsFlyWeight();

            IEnumerable<PropertyInfo> props = importDTOType.GetProperties().ToList().Where(p => p.IsDefined(typeof(ExcelImportAttribute)));
            props.ToList().ForEach(p =>
            {
                string colName = p.GetCustomAttribute<ExcelImportAttribute>().ColName;
                ExcelCol col = excelHeaderRow.Cells.SingleOrDefault(c => c.ColName == colName);
                if (col != null)
                {
                    typeAttrsFlyWeight.PropertyFilterAttrs.Add(
                        new PropertyFilterAttrs
                        {
                            ColIndex = col.ColIndex,
                            PropertyName = p.Name,
                            FilterAttrs = p.GetCustomAttributes<BaseFilterAttribute>()?.ToList()
                        });
                }
            });

            Table[key] = typeAttrsFlyWeight;

            return typeAttrsFlyWeight;
        }

        public List<PropertyFilterAttrs> PropertyFilterAttrs { get; set; } = new List<PropertyFilterAttrs>();
        
    }
}
