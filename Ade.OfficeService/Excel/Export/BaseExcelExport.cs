using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Ade.OfficeService.Excel
{
    public class BaseExcelExport
    {
        private IWorkbook Workbook { get; set; }
        private ISheet Sheet { get; set; }

        internal void Init()
        {
            Workbook = new HSSFWorkbook();
            Sheet = Workbook.CreateSheet();
        }

        public IWorkbook Export(List<BaseExcelExport> baseExportDtos)
        {
            Init();

            ExcelSetHeader();

            ExcelSetDataRow(baseExportDtos);

            Decorate();

            return this.Workbook;
        }

        protected void Decorate()
        {
            DecoratorContext context = new DecoratorContext()
            {
                TypeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(this.GetType())
            };

            GetDecorators().ForEach(d =>
            {
                this.Workbook = d.Decorate(this.Workbook, context);
            });
        }

        /// <summary>
        /// 获取所有的装饰器
        /// </summary>
        /// <returns></returns>
        private List<IDecorator> GetDecorators()
        {
            List<IDecorator> decorators = new List<IDecorator>();
            List<BaseDecorateAttribute> attrs = new List<BaseDecorateAttribute>();
            TypeDecoratorInfo typeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(this.GetType());

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
        protected virtual void ExcelSetHeader()
        {
            IRow row = Sheet.CreateRow(0);
            int colIndex = 0;
            var dict = ExportMappingDictFactory.CreateInstance(this.GetType());
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
        protected virtual void ExcelSetDataRow(List<BaseExcelExport> lst)
        {
            if (lst.Count <= 0)
            {
                return;
            }

            IRow row;
            int colIndex;
            BaseExcelExport dto;

            var dict = ExportMappingDictFactory.CreateInstance(this.GetType());

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
