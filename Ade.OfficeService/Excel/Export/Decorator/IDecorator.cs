using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public interface IDecorator
    {
        IWorkbook Decorate(IWorkbook workbook);
    }
}
