using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService
{
    public interface IExcelImportService
    {
        /// <summary>
        /// 设置数据库是否存在校验的委托方法，如果DTO定义了数据库校验则必须要设置。arg1:tableName, arg2:fieldName
        /// </summary>
        /// <param name="delegateDatabaseExist"></param>
        void SetDelegateDatabaseExist(Func<string, string, bool> delegateDatabaseExist);

        /// <summary>
        /// 校验Excel,返回校验结果
        /// </summary>
        /// <param name="fileUrl"></param>
        /// <returns></returns>
        List<ExcelDataRow> Validate<TTemplate>(string fileUrl)
            where TTemplate : BaseExcelImport;

        /// <summary>
        /// 将ExcelDataRow集合转换为指定类型集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excelDataRows"></param>
        /// <returns></returns>
        List<T> Convert<T>(List<ExcelDataRow> excelDataRows);
    }
}
