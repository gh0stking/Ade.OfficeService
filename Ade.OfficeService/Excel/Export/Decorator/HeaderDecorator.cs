using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using System.Linq;

namespace Ade.OfficeService.Excel
{
    [BindDecorator(typeof(HeaderAttribute))]
    internal class HeaderDecorator : IDecorator
    {
        public IWorkbook Decorate(IWorkbook workbook, DecoratorContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            var attr = (HeaderAttribute)context.TypeDecoratorInfo.TypeDecoratorAttrs.SingleOrDefault(a => a.GetType() == typeof(HeaderAttribute));

            if (attr == null)
            {
                return workbook;
            }
            
            IRow headerRow = workbook.GetSheetAt(0).GetRow(0);
            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();

            font.FontName = attr.FontName;
            font.Color = (short)attr.Color.GetHashCode();
            font.FontHeightInPoints = (short)attr.FontSize;
            if (attr.IsBold)
            {
                font.Boldweight = short.MaxValue;
            }

            style.SetFont(font);

            for (int i = 0; i < headerRow.PhysicalNumberOfCells; i++)
            {
                headerRow.GetCell(i).CellStyle = style;
            }

            return workbook;
        }
    }
}
