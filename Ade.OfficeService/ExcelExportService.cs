using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class ExcelExportService<T>
        where T: IExcelExport
    {
        public static IWorkbook Export(List<T> exportDtos)
        {
            Init();

            ExcelSetHeader();

            ExcelSetDataRow(exportDtos);

            Decorate();

            return Workbook;
        }
        private static IWorkbook Workbook { get; set; }
        private static ISheet Sheet { get; set; }
        private static void Init()
        {
            Workbook = new HSSFWorkbook();
            Sheet = Workbook.CreateSheet();
        }
        private static void Decorate()
        {
            DecoratorContext context = new DecoratorContext()
            {
                TypeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(typeof(T))
            };

            GetDecorators().ForEach(d =>
            {
                Workbook = d.Decorate(Workbook, context);
            });
        }
        /// <summary>
        /// 获取所有的装饰器
        /// </summary>
        /// <returns></returns>
        private static List<IDecorator> GetDecorators()

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
        private static void ExcelSetHeader()
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
        private static void ExcelSetDataRow(List<T> lst)
        {
            if (lst.Count <= 0)
            {
                return;
            }

            IRow row;
            int colIndex;
            IExcelExport dto;

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
