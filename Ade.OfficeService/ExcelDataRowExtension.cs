using NPOI.SS.UserModel;
using System;
using System.Text.RegularExpressions;

namespace Ade.OfficeService.Excel
{
    public static class ExcelDataRowExtension
    {
        public static void SetState(this ExcelDataRow row, bool isValid, ExcelCol col,string errorMsg)
        {
            if (!isValid)
            {
                row.IsValid = isValid;
                row.ErrorMsgs.Add(col, errorMsg);
            }
        }
    }
}
