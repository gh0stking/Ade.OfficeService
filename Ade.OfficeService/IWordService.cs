using Ade.OfficeService.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService
{
    public interface IWordService
    {
        /// <summary>
        /// 得到渲染后的Word字节数组
        /// </summary>
        /// <typeparam name="TTemplate"></typeparam>
        /// <typeparam name="TIn"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        byte[] ExportByTemplate<TTemplate, TIn>(IEnumerable<TIn> data);
    }
}
