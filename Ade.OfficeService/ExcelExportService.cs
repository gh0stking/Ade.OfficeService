using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class ExcelExportService
    {
        public static IWorkbook Export<T>(List<T> exportDtos)
            where T:class,new()
        {
            Init();

            ExcelSetHeader<T>();

            ExcelSetDataRow(exportDtos);

            Decorate<T>();

            return Workbook;
        }
        private static IWorkbook Workbook { get; set; }
        private static ISheet Sheet { get; set; }
        private static void Init()
        {
            Workbook = new HSSFWorkbook();
            Sheet = Workbook.CreateSheet();
        }
        private static void Decorate<T>()
              where T : class, new()
        {
            DecoratorContext context = new DecoratorContext()
            {
                TypeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(typeof(T))
            };

            GetDecorators<T>().ForEach(d =>
            {
                Workbook = d.Decorate(Workbook, context);
            });
        }
        /// <summary>
        /// 获取所有的装饰器
        /// </summary>
        /// <returns></returns>
        private static List<IDecorator> GetDecorators<T>()
                where T : class, new()

        {
            List<IDecorator> decorators = new List<IDecorator>();
            List<BaseDecorateAttribute> attrs = new List<BaseDecorateAttribute>();
            TypeDecoratorInfo typeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(typeof(T));

            attrs.AddRange(typeDecoratorInfo.TypeDecoratorAttrs);
            typeDecoratorInfo.PropertyDecoratorInfos.ForEach(a => attrs.AddRange(a.DecoratorAttrs));

            attrs.Distinct(new DecoratorAttributeComparer()).ToList().ForEach
            (a =>
            {
                var decorator = DecoratorFactory.CreateInstance(a.GetType());
                if (decorator != null)
                {
                    decorators.Add(decorator);
                }
            });

            return decorators;
        }
        /// <summary>
        /// 设置表头
        /// </summary>
        private static void ExcelSetHeader<T>()
            where T:class,new()
        {
            IRow row = Sheet.CreateRow(0);
            int colIndex = 0;
            var dict = ExportMappingDictFactory.CreateInstance(typeof(T));
            foreach (var kvp in dict)
            {
                row.CreateCell(colIndex).SetCellValue(kvp.Value);
                colIndex++;
            }
        }
        /// <summary>
        /// 设置数据
        /// </summary>
        /// <param name="lst"></param>
        private static void ExcelSetDataRow<T>(List<T> lst)
                where T : class, new()
        {
            if (lst.Count <= 0)
            {
                return;
            }

            IRow row;
            int colIndex;
            T dto;

            var dict = ExportMappingDictFactory.CreateInstance(typeof(T));

            for (int i = 0; i < lst.Count; i++)
            {
                colIndex = 0;
                row = Sheet.CreateRow(i + 1);
                dto = lst[i];

                foreach (var kvp in dict)
                {
                    row.CreateCell(colIndex).SetCellValue(dto.GetStringValue(kvp.Key));
                    colIndex++;
                }
            }
        }
    }
}
