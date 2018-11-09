using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService
{
    public interface IExcelExportService
    {
        /// <summary>
        /// 根据模板定义，将数据集合导出为Excel文件的byte数组
        /// </summary>
        /// <typeparam name="TTemplate"></typeparam>
        /// <typeparam name="TIn"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        byte[] Export<TTemplate, TIn>(IEnumerable<TIn> dataCollection);
    }
}
