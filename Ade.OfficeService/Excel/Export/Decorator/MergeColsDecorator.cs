using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using System.Linq;
using NPOI.SS.Util;

namespace Ade.OfficeService.Excel
{
    [BindDecorator(typeof(MergeColsAttribute))]
    public class MergeColsDecorator : IDecorator
    {
        public IWorkbook Decorate(IWorkbook workbook, DecoratorContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            var propertyDecoratorInfos = context.TypeDecoratorInfo.PropertyDecoratorInfos;
            ISheet sheet = workbook.GetSheetAt(0);
            foreach (var item in propertyDecoratorInfos)
            {
                if (item.DecoratorAttrs.SingleOrDefault(a => a.GetType() == typeof(MergeColsAttribute)) != null)
                {
                    InternalDecorate(sheet, item.ColIndex);
                }
            }

            return workbook;
        }

        private void InternalDecorate(ISheet sheet, int colIndex)
        {
            string currentCellValue;
            int startRowIndex = 1;
            CellRangeAddress mergeRangeAddress;
            string startCellValue = sheet.GetRow(1).GetCell(colIndex).StringCellValue;

            for (int rowIndex = 1; rowIndex < sheet.PhysicalNumberOfRows; rowIndex++)
            {
                currentCellValue = sheet.GetRow(rowIndex).GetCell(colIndex).StringCellValue;

                if (currentCellValue.Trim() != startCellValue.Trim())
                {
                    mergeRangeAddress = new CellRangeAddress(startRowIndex, rowIndex - 1, colIndex, colIndex);
                    sheet.AddMergedRegion(mergeRangeAddress);

                    sheet.GetRow(startRowIndex).GetCell(colIndex).CellStyle.VerticalAlignment = VerticalAlignment.Center;

                    startRowIndex = rowIndex;
                    startCellValue = currentCellValue;
                }

                if (rowIndex == sheet.PhysicalNumberOfRows - 1 && startRowIndex != rowIndex)
                {
                    mergeRangeAddress = new CellRangeAddress(startRowIndex, rowIndex, colIndex, colIndex);
                    sheet.AddMergedRegion(mergeRangeAddress);

                    sheet.GetRow(startRowIndex).GetCell(colIndex).CellStyle.VerticalAlignment = VerticalAlignment.Center;
                }
            }
        }
    }
}
