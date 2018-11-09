//using System;
//using System.Collections;
//using System.Collections.Generic;
//using System.Text;
//using System.Linq;
//using System.Reflection;

//namespace Ade.OfficeService.Excel
//{
//    public class TypeDecorateInfoFactory
//    {
//        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));
//        public static TypeFilterInfo CreateInstance(Type exportDTOType)
//        {
//            if (exportDTOType == null)
//            {
//                throw new ArgumentNullException("exportDtoType");
//            }
          
//            var key = exportDTOType;
//            if (Table[key] != null)
//            {
//                return (TypeFilterInfo)Table[key];
//            }

//            TypeDecorateInfo typeFilterInfo = new TypeDecorateInfo()
//            {
//                GlobalDecorators = new List<IDecorator>(),
//                PropertyExportInfos = new List<PropertyDecorateInfo>() { }
//            };

//            IEnumerable<PropertyInfo> props = exportDTOType.GetProperties().ToList().Where(p => p.IsDefined(typeof(ExcelExportAttribute)));
//            props.ToList().ForEach(p =>
//            {
//                string colName = p.GetCustomAttribute<ExcelExportAttribute>().ColName;

//                p.GetCustomAttributes(typeof(BaseDecorateAttribute)).

//                ExcelCol col = excelHeaderRow.Cells.SingleOrDefault(c => c.ColName == colName);
//                if (col != null)
//                {
//                    typeFilterInfo.PropertyFilterInfos.Add(
//                        new PropertyFilterInfo
//                        {
//                            ColIndex = col.ColIndex,
//                            FilterAttrs = p.GetCustomAttributes<BaseFilterAttribute>()?.ToList()
//                        });
//                }
//            });

//            Table[key] = typeFilterInfo;

//            return typeFilterInfo;
//        }
//    }
//}
