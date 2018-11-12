using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    [BindDecorator(typeof(WrapTextAttribute))]
    internal class WrapTextDecorator : IDecorator
    {
        public IWorkbook Decorate(IWorkbook workbook, DecoratorContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            var attr = context.TypeDecoratorInfo.GetDecorateAttr<WrapTextAttribute>();

            if (attr == null)
            {
                return workbook;
            }

            ISheet sheet = workbook.GetSheetAt(0);
            IRow row;
            if (sheet.PhysicalNumberOfRows > 0)
            {
                for (int i = 0; i < sheet.PhysicalNumberOfRows; i++)
                {
                    row = sheet.GetRow(i);
                    for (int colIndex = 0; colIndex < row.PhysicalNumberOfCells; colIndex++)
                    {
                        row.GetCell(colIndex).CellStyle.WrapText = true;
                    }
                }
            }

            return workbook;
        }
    }
}
