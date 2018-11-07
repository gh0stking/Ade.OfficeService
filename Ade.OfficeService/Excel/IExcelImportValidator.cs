using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public interface IFilter
    {
        /// <summary>
        /// 注册校验器
        /// </summary>
        /// <param name="importDTOType"></param>
        //void Register(Type importDTOType, ExcelHeaderRow headerRow, List<ExcelDataRow> excelDataRows, string errorMsg, Func<string, string, bool> delegateDatabaseValidate = null);


        /// <summary>
        /// 校验
        /// </summary>
        /// <param name="rowWrappers"></param>
        /// <returns></returns>
        List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows);
    }

}
